/*
  Technical Standard

  Menu command handlers for File, Edit, Format, and View menu operations.
  Bridges user actions to core functionality modules.
*/

#include "commands.h"
#include "core/file_dialog_filters.h"
#include "core/globals.h"
#include "editor.h"
#include "file.h"
#include "ui.h"
#include "settings.h"
#include "tab_layout.h"
#include "resource.h"
#include "lang/lang.h"
#include <commdlg.h>
#include <richedit.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <algorithm>
#include <string>
#include <vector>
#include <array>
#include <sstream>

namespace
{
bool SelectionEffectActive(DWORD effectMask, DWORD effectFlag)
{
    CHARFORMAT2W current{};
    current.cbSize = sizeof(current);
    SendMessageW(g_hwndEditor, EM_GETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&current));
    if ((current.dwMask & effectMask) == 0)
        return false;
    return (current.dwEffects & effectFlag) != 0;
}

void ToggleSelectionEffect(DWORD effectMask, DWORD effectFlag)
{
    if (!g_hwndEditor || g_state.largeFileMode)
        return;

    DWORD selectionStart = 0;
    DWORD selectionEnd = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
    const bool hasSelection = selectionStart != selectionEnd;

    CHARFORMAT2W apply{};
    apply.cbSize = sizeof(apply);
    apply.dwMask = effectMask;
    apply.dwEffects = SelectionEffectActive(effectMask, effectFlag) ? 0 : effectFlag;
    const LRESULT ok = SendMessageW(g_hwndEditor, EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&apply));
    if (ok != 0 && hasSelection && !g_state.modified)
    {
        g_state.modified = true;
        UpdateTitle();
    }
    SetFocus(g_hwndEditor);
}

#if !defined(TECHNICAL_STANDARD_NOTE_SPLIT_COMMAND_MODULES)
using fnPickIconDlg = BOOL(WINAPI *)(HWND, LPWSTR, UINT, int *);

bool IsIconResourceContainer(const std::wstring &path)
{
    const wchar_t *ext = PathFindExtensionW(path.c_str());
    if (!ext || ext[0] == L'\0')
        return false;
    return lstrcmpiW(ext, L".dll") == 0 ||
           lstrcmpiW(ext, L".exe") == 0 ||
           lstrcmpiW(ext, L".icl") == 0 ||
           lstrcmpiW(ext, L".mun") == 0;
}

std::wstring DefaultSystemIconLibrary()
{
    wchar_t path[MAX_PATH] = {};
    DWORD len = ExpandEnvironmentStringsW(L"%SystemRoot%\\SystemResources\\imageres.dll.mun", path, MAX_PATH);
    if (len == 0 || len >= MAX_PATH || GetFileAttributesW(path) == INVALID_FILE_ATTRIBUTES)
    {
        len = ExpandEnvironmentStringsW(L"%SystemRoot%\\System32\\shell32.dll", path, MAX_PATH);
        if (len == 0 || len >= MAX_PATH)
            return L"shell32.dll";
    }
    return path;
}

HICON LoadIconFromPath(const std::wstring &path, int iconIndex)
{
    if (path.empty())
        return nullptr;

    const bool fileIconOnly = !IsIconResourceContainer(path);
    if (fileIconOnly)
    {
        return static_cast<HICON>(LoadImageW(nullptr, path.c_str(), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE));
    }

    HICON hLarge = nullptr;
    HICON hSmall = nullptr;
    const UINT extracted = ExtractIconExW(path.c_str(), iconIndex, &hLarge, &hSmall, 1);
    if (extracted == 0)
        return nullptr;

    HICON chosen = hLarge ? hLarge : hSmall;
    if (!chosen)
        return nullptr;

    if (hLarge && hLarge != chosen)
        DestroyIcon(hLarge);
    if (hSmall && hSmall != chosen)
        DestroyIcon(hSmall);
    return chosen;
}

bool OpenSystemIconPicker(HWND owner, LPWSTR iconPath, UINT cchIconPath, int *iconIndex)
{
    HMODULE hShell32 = LoadLibraryW(L"shell32.dll");
    if (!hShell32)
        return false;

#ifdef __GNUC__
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wcast-function-type"
#endif
    auto pickIconDlg = reinterpret_cast<fnPickIconDlg>(GetProcAddress(hShell32, "PickIconDlg"));
#ifdef __GNUC__
#pragma GCC diagnostic pop
#endif
    if (!pickIconDlg)
    {
        FreeLibrary(hShell32);
        return false;
    }

    BOOL ok = pickIconDlg(owner, iconPath, cchIconPath, iconIndex);
    FreeLibrary(hShell32);
    return ok != FALSE;
}

struct PerfBenchmarkResult
{
    std::wstring label;
    size_t bytes = 0;
    size_t operations = 0;
    double readMs = 0.0;
    double detectMs = 0.0;
    double decodeMs = 0.0;
    double totalMs = 0.0;
    double budgetMs = 0.0;
    size_t charCount = 0;
    std::wstring notes;
    bool passed = false;
    bool success = false;
};

bool WriteAllToHandle(HANDLE handle, const void *buffer, DWORD bytes)
{
    const BYTE *ptr = static_cast<const BYTE *>(buffer);
    DWORD remaining = bytes;
    while (remaining > 0)
    {
        DWORD written = 0;
        if (!WriteFile(handle, ptr, remaining, &written, nullptr))
            return false;
        if (written == 0)
            return false;
        ptr += written;
        remaining -= written;
    }
    return true;
}

double ElapsedMs(const LARGE_INTEGER &start, const LARGE_INTEGER &end, const LARGE_INTEGER &freq)
{
    if (freq.QuadPart <= 0)
        return 0.0;
    return static_cast<double>(end.QuadPart - start.QuadPart) * 1000.0 / static_cast<double>(freq.QuadPart);
}

std::wstring BenchmarkDirectoryPath()
{
    wchar_t localAppData[MAX_PATH] = {};
    DWORD len = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    std::wstring root;
    if (len > 0 && len < MAX_PATH)
    {
        root = localAppData;
    }
    else
    {
        wchar_t tempPath[MAX_PATH] = {};
        DWORD tempLen = GetTempPathW(MAX_PATH, tempPath);
        if (tempLen > 0 && tempLen < MAX_PATH)
            root = tempPath;
        else
            root = L".";
    }

    root += L"\\TechnicalStandardNote";
    CreateDirectoryW(root.c_str(), nullptr);
    root += L"\\benchmarks";
    CreateDirectoryW(root.c_str(), nullptr);
    return root;
}

std::wstring BenchmarkTimestamp()
{
    SYSTEMTIME st{};
    GetLocalTime(&st);
    wchar_t stamp[64] = {};
    wsprintfW(stamp, L"%04u%02u%02u-%02u%02u%02u",
              st.wYear, st.wMonth, st.wDay,
              st.wHour, st.wMinute, st.wSecond);
    return stamp;
}

bool CreateBenchmarkFile(const std::wstring &path, size_t targetBytes)
{
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE)
        return false;

    static constexpr char kLine[] =
        "TechnicalStandardNote benchmark line 0123456789 abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ.\r\n";
    const DWORD lineBytes = static_cast<DWORD>(sizeof(kLine) - 1);
    size_t writtenTotal = 0;
    bool ok = true;

    while (writtenTotal < targetBytes)
    {
        DWORD chunk = lineBytes;
        const size_t remaining = targetBytes - writtenTotal;
        if (remaining < lineBytes)
            chunk = static_cast<DWORD>(remaining);
        if (!WriteAllToHandle(file, kLine, chunk))
        {
            ok = false;
            break;
        }
        writtenTotal += chunk;
    }

    CloseHandle(file);
    return ok;
}

HWND CreateHiddenBenchmarkEditor()
{
    if (!g_hwndMain)
        return nullptr;

    const wchar_t *editorClass = g_editorClassName.empty() ? MSFTEDIT_CLASS : g_editorClassName.c_str();
    HWND hwnd = CreateWindowExW(0, editorClass, nullptr,
                                WS_CHILD | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN | ES_NOHIDESEL,
                                0, 0, 640, 480,
                                g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!hwnd)
        return nullptr;

    ConfigureEditorControl(hwnd);
    SendMessageW(hwnd, EM_EXLIMITTEXT, 0, static_cast<LPARAM>(-1));
    HFONT hFont = g_state.hFont ? g_state.hFont : TabGetRegularFont();
    if (hFont)
        SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), FALSE);
    return hwnd;
}

bool SetBenchmarkEditorText(HWND hwnd, const std::wstring &text)
{
    if (!hwnd)
        return false;
    return SendMessageW(hwnd, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(text.c_str())) != FALSE;
}

std::wstring BuildScrollBenchmarkText(size_t lines)
{
    static constexpr wchar_t kLine[] = L"TechnicalStandardNote scroll benchmark line 0123456789 abcdefghijklmnopqrstuvwxyz\r\n";
    static constexpr size_t kLineChars = (sizeof(kLine) / sizeof(wchar_t)) - 1;
    std::wstring text;
    text.reserve(lines * kLineChars);
    for (size_t i = 0; i < lines; ++i)
        text += kLine;
    return text;
}

bool RunOpenBenchmark(const std::wstring &filePath, const std::wstring &label, double budgetMs, PerfBenchmarkResult &outResult)
{
    outResult = {};
    outResult.label = label;
    outResult.budgetMs = budgetMs;

    LARGE_INTEGER freq{};
    QueryPerformanceFrequency(&freq);

    HANDLE file = CreateFileW(filePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE)
        return false;

    LARGE_INTEGER sizeLarge{};
    if (!GetFileSizeEx(file, &sizeLarge) || sizeLarge.QuadPart < 0 || sizeLarge.QuadPart > static_cast<LONGLONG>(64 * 1024 * 1024))
    {
        CloseHandle(file);
        return false;
    }

    outResult.bytes = static_cast<size_t>(sizeLarge.QuadPart);
    std::vector<BYTE> data(outResult.bytes);

    LARGE_INTEGER readStart{}, readEnd{};
    QueryPerformanceCounter(&readStart);
    DWORD bytesRead = 0;
    bool readOk = true;
    if (!data.empty())
    {
        readOk = ReadFile(file, data.data(), static_cast<DWORD>(data.size()), &bytesRead, nullptr) != FALSE;
        readOk = readOk && (bytesRead == data.size());
    }
    QueryPerformanceCounter(&readEnd);
    CloseHandle(file);
    if (!readOk)
        return false;
    outResult.readMs = ElapsedMs(readStart, readEnd, freq);

    LARGE_INTEGER detectStart{}, detectEnd{};
    QueryPerformanceCounter(&detectStart);
    const auto detected = DetectEncoding(data);
    const Encoding encoding = detected.first;
    QueryPerformanceCounter(&detectEnd);
    outResult.detectMs = ElapsedMs(detectStart, detectEnd, freq);

    LARGE_INTEGER decodeStart{}, decodeEnd{};
    QueryPerformanceCounter(&decodeStart);
    std::wstring decoded = DecodeText(data, encoding);
    QueryPerformanceCounter(&decodeEnd);
    outResult.decodeMs = ElapsedMs(decodeStart, decodeEnd, freq);

    outResult.charCount = decoded.size();
    outResult.totalMs = outResult.readMs + outResult.detectMs + outResult.decodeMs;
    outResult.notes = L"open pipeline";
    outResult.passed = outResult.totalMs <= outResult.budgetMs;
    outResult.success = true;
    return true;
}

bool RunTypingBurstBenchmark(const std::wstring &label, double budgetMs, PerfBenchmarkResult &outResult)
{
    outResult = {};
    outResult.label = label;
    outResult.budgetMs = budgetMs;
    outResult.notes = L"EM_REPLACESEL typing loop";

    HWND hwnd = CreateHiddenBenchmarkEditor();
    if (!hwnd)
        return false;

    static constexpr wchar_t kChunk[] = L"technical-standard-note typing burst 0123456789 abcdefghijklmnopqrstuvwxyz\r\n";
    static constexpr size_t kChunkChars = (sizeof(kChunk) / sizeof(wchar_t)) - 1;
    static constexpr int kIterations = 1200;

    LARGE_INTEGER freq{};
    QueryPerformanceFrequency(&freq);
    LARGE_INTEGER start{}, end{};
    QueryPerformanceCounter(&start);

    SendMessageW(hwnd, EM_SETSEL, 0, -1);
    SendMessageW(hwnd, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(L""));
    for (int i = 0; i < kIterations; ++i)
    {
        SendMessageW(hwnd, EM_SETSEL, static_cast<WPARAM>(-1), static_cast<LPARAM>(-1));
        SendMessageW(hwnd, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(kChunk));
    }

    QueryPerformanceCounter(&end);
    DestroyWindow(hwnd);

    outResult.totalMs = ElapsedMs(start, end, freq);
    outResult.operations = static_cast<size_t>(kIterations);
    outResult.charCount = static_cast<size_t>(kIterations) * kChunkChars;
    outResult.bytes = outResult.charCount * sizeof(wchar_t);
    outResult.passed = outResult.totalMs <= outResult.budgetMs;
    outResult.success = true;
    return true;
}

bool RunScrollStressBenchmark(const std::wstring &label, double budgetMs, PerfBenchmarkResult &outResult)
{
    outResult = {};
    outResult.label = label;
    outResult.budgetMs = budgetMs;
    outResult.notes = L"EM_LINESCROLL stress loop";

    HWND hwnd = CreateHiddenBenchmarkEditor();
    if (!hwnd)
        return false;

    static constexpr size_t kLines = 30000;
    static constexpr int kScrollOps = 3000;

    const std::wstring text = BuildScrollBenchmarkText(kLines);
    if (!SetBenchmarkEditorText(hwnd, text))
    {
        DestroyWindow(hwnd);
        return false;
    }

    LARGE_INTEGER freq{};
    QueryPerformanceFrequency(&freq);
    LARGE_INTEGER start{}, end{};

    SendMessageW(hwnd, EM_SETSEL, 0, 0);
    QueryPerformanceCounter(&start);
    for (int i = 0; i < kScrollOps; ++i)
        SendMessageW(hwnd, EM_LINESCROLL, 0, 1);
    QueryPerformanceCounter(&end);

    DestroyWindow(hwnd);

    outResult.totalMs = ElapsedMs(start, end, freq);
    outResult.operations = static_cast<size_t>(kScrollOps);
    outResult.charCount = text.size();
    outResult.bytes = text.size() * sizeof(wchar_t);
    outResult.passed = outResult.totalMs <= outResult.budgetMs;
    outResult.success = true;
    return true;
}

std::wstring FormatBenchmarkReport(const std::vector<PerfBenchmarkResult> &results)
{
    std::wostringstream ss;
    ss << L"Technical Standard Note Performance Benchmark\n";
    ss << L"Scope: open pipeline + typing burst + scroll stress\n";
    ss << L"Hardware dependent. Use this for regression tracking.\n\n";

    for (const PerfBenchmarkResult &result : results)
    {
        ss << L"Case: " << result.label << L"\n";
        if (!result.success)
        {
            ss << L"Status: FAIL (benchmark case could not run)\n\n";
            continue;
        }

        if (result.bytes > 0)
        {
            const double sizeMb = static_cast<double>(result.bytes) / (1024.0 * 1024.0);
            ss << L"Size: " << sizeMb << L" MB\n";
            if (result.readMs > 0.0 || result.detectMs > 0.0 || result.decodeMs > 0.0)
            {
                const double throughput = (result.totalMs > 0.0) ? (sizeMb * 1000.0 / result.totalMs) : 0.0;
                ss << L"Read: " << result.readMs << L" ms\n";
                ss << L"Detect: " << result.detectMs << L" ms\n";
                ss << L"Decode: " << result.decodeMs << L" ms\n";
                ss << L"Throughput: " << throughput << L" MB/s\n";
            }
        }
        if (result.operations > 0)
        {
            const double opsPerSec = (result.totalMs > 0.0) ? (static_cast<double>(result.operations) * 1000.0 / result.totalMs) : 0.0;
            ss << L"Operations: " << result.operations << L"\n";
            ss << L"Ops/sec: " << opsPerSec << L"\n";
        }
        if (result.charCount > 0)
            ss << L"Chars: " << result.charCount << L"\n";
        ss << L"Total: " << result.totalMs << L" ms (budget " << result.budgetMs << L" ms)\n";
        if (!result.notes.empty())
            ss << L"Notes: " << result.notes << L"\n";
        ss << L"Status: " << (result.passed ? L"PASS" : L"WARN") << L"\n\n";
    }

    return ss.str();
}

std::string WideToUtf8(const std::wstring &text)
{
    if (text.empty())
        return {};
    const int bytes = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (bytes <= 0)
        return {};
    std::string out(bytes, '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), out.data(), bytes, nullptr, nullptr);
    return out;
}

bool SaveBenchmarkReport(const std::wstring &path, const std::wstring &content)
{
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE)
        return false;

    static const BYTE utf8Bom[] = {0xEF, 0xBB, 0xBF};
    bool ok = WriteAllToHandle(file, utf8Bom, sizeof(utf8Bom));
    const std::string body = WideToUtf8(content);
    if (ok && !body.empty())
        ok = WriteAllToHandle(file, body.data(), static_cast<DWORD>(body.size()));
    CloseHandle(file);
    return ok;
}
#endif
}

bool ConfirmDiscard()
{
    if (!g_state.modified)
        return true;
    if (g_state.filePath.empty())
    {
        std::wstring text = GetEditorText();
        if (text.empty())
            return true;
    }
    const auto &lang = GetLangStrings();
    std::wstring filename = g_state.filePath.empty() ? lang.untitled : PathFindFileNameW(g_state.filePath.c_str());
    std::wstring msg;
    msg.reserve(lang.msgSaveChanges.size() + filename.size() + 2);
    msg = lang.msgSaveChanges;
    msg += filename;
    msg += L"?";
    int result = MessageBoxW(g_hwndMain, msg.c_str(), lang.appName.c_str(), MB_YESNOCANCEL | MB_ICONWARNING);
    if (result == IDYES)
    {
        FileSave();
        return !g_state.modified;
    }
    return result == IDNO;
}

void FileNew()
{
    if (!ConfirmDiscard())
        return;
    const bool wasLargeMode = g_state.largeFileMode;
    g_state.largeFileMode = false;
    g_state.largeFileBytes = 0;
    if (wasLargeMode)
        ApplyWordWrap();
    SetEditorText(L"");
    g_state.filePath.clear();
    g_state.modified = false;
    g_state.encoding = Encoding::UTF8;
    g_state.lineEnding = LineEnding::CRLF;
    UpdateTitle();
    UpdateStatus();
}

void FileOpen()
{
    if (!ConfirmDiscard())
        return;
    const auto &lang = GetLangStrings();
    const std::wstring filter = BuildTextDocumentsFilter(lang);
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_ENABLESIZING;
    if (GetOpenFileNameW(&ofn))
        LoadFile(path);
}

void FileSave()
{
    if (g_state.filePath.empty())
        FileSaveAs();
    else
        SaveToPath(g_state.filePath);
}

void FileSaveAs()
{
    const auto &lang = GetLangStrings();
    const std::wstring filter = BuildTextDocumentsFilter(lang);
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrDefExt = L"txt";
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_ENABLESIZING;
    if (GetSaveFileNameW(&ofn))
        SaveToPath(path);
}

void FilePrint()
{
    std::wstring text = GetEditorText();
    PRINTDLGW pd = {};
    pd.lStructSize = sizeof(pd);
    pd.hwndOwner = g_hwndMain;
    pd.Flags = PD_RETURNDC | PD_NOPAGENUMS | PD_NOSELECTION;
    if (!PrintDlgW(&pd))
        return;
    HDC hDC = pd.hDC;
    DOCINFOW di = {};
    di.cbSize = sizeof(di);
    const auto &lang = GetLangStrings();
    std::wstring docName = g_state.filePath.empty() ? lang.untitled : PathFindFileNameW(g_state.filePath.c_str());
    di.lpszDocName = docName.c_str();
    if (StartDocW(hDC, &di) > 0)
    {
        int pageWidth = GetDeviceCaps(hDC, HORZRES);
        int pageHeight = GetDeviceCaps(hDC, VERTRES);
        int marginX = pageWidth / 10, marginY = pageHeight / 10;
        int printWidth = pageWidth - 2 * marginX;
        int printHeight = pageHeight - 2 * marginY;
        HFONT hPrintFont = CreateFontW(-MulDiv(10, GetDeviceCaps(hDC, LOGPIXELSY), 72),
                                       0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                       OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY,
                                       FIXED_PITCH | FF_MODERN, g_state.fontName.c_str());
        HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(hDC, hPrintFont));
        TEXTMETRICW tm;
        GetTextMetricsW(hDC, &tm);
        int lineHeight = tm.tmHeight + tm.tmExternalLeading;
        int linesPerPage = printHeight / lineHeight;
        std::vector<std::wstring> lines;
        std::wstring line;
        for (size_t i = 0; i <= text.size(); ++i)
        {
            if (i == text.size() || text[i] == L'\n' || text[i] == L'\r')
            {
                lines.push_back(line);
                line.clear();
                if (i < text.size() && text[i] == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n')
                    ++i;
            }
            else
                line += text[i];
        }
        int totalLines = static_cast<int>(lines.size());
        int lineIndex = 0;
        while (lineIndex < totalLines)
        {
            StartPage(hDC);
            int y = marginY;
            for (int i = 0; i < linesPerPage && lineIndex < totalLines; ++i, ++lineIndex)
            {
                RECT rc = {marginX, y, marginX + printWidth, y + lineHeight};
                DrawTextW(hDC, lines[lineIndex].c_str(), -1, &rc, DT_LEFT | DT_NOPREFIX | DT_SINGLELINE);
                y += lineHeight;
            }
            EndPage(hDC);
        }
        SelectObject(hDC, hOldFont);
        DeleteObject(hPrintFont);
        EndDoc(hDC);
    }
    DeleteDC(hDC);
}

void FilePageSetup()
{
    g_pageSetup.lStructSize = sizeof(g_pageSetup);
    g_pageSetup.hwndOwner = g_hwndMain;
    g_pageSetup.Flags = PSD_MARGINS | PSD_INHUNDREDTHSOFMILLIMETERS;
    PageSetupDlgW(&g_pageSetup);
}

void EditUndo() { SendMessageW(g_hwndEditor, EM_UNDO, 0, 0); }
void EditRedo() { SendMessageW(g_hwndEditor, EM_REDO, 0, 0); }
void EditCut() { SendMessageW(g_hwndEditor, WM_CUT, 0, 0); }
void EditCopy() { SendMessageW(g_hwndEditor, WM_COPY, 0, 0); }
void EditPaste() { SendMessageW(g_hwndEditor, WM_PASTE, 0, 0); }
void EditDelete()
{
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    if (start != end)
    {
        SendMessageW(g_hwndEditor, WM_CLEAR, 0, 0);
        return;
    }
    SendMessageW(g_hwndEditor, WM_KEYDOWN, VK_DELETE, 1);
    SendMessageW(g_hwndEditor, WM_KEYUP, VK_DELETE, (1u << 30) | (1u << 31));
}
void EditSelectAll() { SendMessageW(g_hwndEditor, EM_SETSEL, 0, -1); }

void EditTimeDate()
{
    SYSTEMTIME st{};
    GetLocalTime(&st);

    wchar_t timeBuf[64] = {};
    wchar_t dateBuf[64] = {};
    wchar_t out[160] = {};

    const int timeLen = GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, 0, &st, nullptr, timeBuf, static_cast<int>(sizeof(timeBuf) / sizeof(timeBuf[0])));
    const int dateLen = GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &st, nullptr, dateBuf, static_cast<int>(sizeof(dateBuf) / sizeof(dateBuf[0])), nullptr);
    if (timeLen > 0 && dateLen > 0)
    {
        wsprintfW(out, L"%s %s", timeBuf, dateBuf);
    }
    else
    {
        wsprintfW(out, L"%02d:%02d %02d/%02d/%04d", st.wHour, st.wMinute, st.wMonth, st.wDay, st.wYear);
    }

    SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(out));
}

void FormatWordWrap()
{
    g_state.wordWrap = !g_state.wordWrap;
    CheckMenuItem(GetMenu(g_hwndMain), IDM_FORMAT_WORDWRAP, g_state.wordWrap ? MF_CHECKED : MF_UNCHECKED);
    ApplyWordWrap();
    SaveFontSettings();
}

void FormatBold()
{
    ToggleSelectionEffect(CFM_BOLD, CFE_BOLD);
}

void FormatItalic()
{
    ToggleSelectionEffect(CFM_ITALIC, CFE_ITALIC);
}

void FormatStrikethrough()
{
    ToggleSelectionEffect(CFM_STRIKEOUT, CFE_STRIKEOUT);
}

void ViewZoomIn()
{
    static const int levels[] = {25, 50, 75, 100, 125, 150, 175, 200, 250, 300, 350, 400, 450, 500};
    for (int l : levels)
    {
        if (l > g_state.zoomLevel)
        {
            g_state.zoomLevel = l;
            ApplyZoom();
            UpdateStatus();
            SaveFontSettings();
            return;
        }
    }
}

void ViewZoomOut()
{
    static const int levels[] = {25, 50, 75, 100, 125, 150, 175, 200, 250, 300, 350, 400, 450, 500};
    for (int i = 13; i >= 0; --i)
    {
        if (levels[i] < g_state.zoomLevel)
        {
            g_state.zoomLevel = levels[i];
            ApplyZoom();
            UpdateStatus();
            SaveFontSettings();
            return;
        }
    }
}

void ViewZoomDefault()
{
    g_state.zoomLevel = ZOOM_DEFAULT;
    ApplyZoom();
    UpdateStatus();
    SaveFontSettings();
}

void ViewStatusBar()
{
    g_state.showStatusBar = !g_state.showStatusBar;
    CheckMenuItem(GetMenu(g_hwndMain), IDM_VIEW_STATUSBAR, g_state.showStatusBar ? MF_CHECKED : MF_UNCHECKED);
    ResizeControls();
    UpdateStatus();
    SaveFontSettings();
}

#if !defined(TECHNICAL_STANDARD_NOTE_SPLIT_COMMAND_MODULES)
void ViewChangeIcon()
{
    const auto &lang = GetLangStrings();
    const std::wstring filter = BuildIconFilesFilter(lang);
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
    if (GetOpenFileNameW(&ofn))
        ApplyCustomIcon(path, 0);
}

void ViewChooseSystemIcon()
{
    std::wstring initialPath = DefaultSystemIconLibrary();
    int iconIndex = 0;

    if (!g_state.customIconPath.empty() && IsIconResourceContainer(g_state.customIconPath))
    {
        initialPath = g_state.customIconPath;
        iconIndex = g_state.customIconIndex;
    }

    wchar_t iconPath[MAX_PATH] = {};
    wcsncpy_s(iconPath, initialPath.c_str(), _TRUNCATE);

    if (!OpenSystemIconPicker(g_hwndMain, iconPath, MAX_PATH, &iconIndex))
        return;

    ApplyCustomIcon(iconPath, iconIndex);
}

void ViewResetIcon()
{
    if (g_hCustomIcon)
    {
        DestroyIcon(g_hCustomIcon);
        g_hCustomIcon = nullptr;
    }
    g_state.customIconPath.clear();
    g_state.customIconIndex = 0;
    HICON hDefaultIcon = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDI_NOTEPAD));
    SendMessageW(g_hwndMain, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(hDefaultIcon));
    SendMessageW(g_hwndMain, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(hDefaultIcon));
    SaveFontSettings();
}

bool ApplyCustomIcon(const std::wstring &iconPath, int iconIndex, bool showError)
{
    const auto &lang = GetLangStrings();
    HICON hNewIcon = LoadIconFromPath(iconPath, iconIndex);
    if (!hNewIcon)
    {
        if (showError)
            MessageBoxW(g_hwndMain, lang.msgIconLoadFailed.c_str(), lang.appName.c_str(), MB_ICONERROR);
        return false;
    }

    if (g_hCustomIcon && g_hCustomIcon != hNewIcon)
        DestroyIcon(g_hCustomIcon);

    g_hCustomIcon = hNewIcon;
    g_state.customIconPath = iconPath;
    g_state.customIconIndex = iconIndex;
    SendMessageW(g_hwndMain, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(hNewIcon));
    SendMessageW(g_hwndMain, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(hNewIcon));
    SaveFontSettings();
    return true;
}

bool RunPerformanceBenchmark(bool interactive, std::wstring *outReportPath, bool *outAllPassed, bool *outAllExecuted)
{
    const HCURSOR oldCursor = interactive ? SetCursor(LoadCursorW(nullptr, IDC_WAIT)) : nullptr;

    const std::wstring benchDir = BenchmarkDirectoryPath();
    struct BenchmarkCase
    {
        size_t bytes;
        const wchar_t *label;
        double budgetMs;
    };

    const std::array<BenchmarkCase, 3> cases = {{
        {1u * 1024u * 1024u, L"Open 1 MB", 35.0},
        {5u * 1024u * 1024u, L"Open 5 MB", 150.0},
        {20u * 1024u * 1024u, L"Open 20 MB", 650.0},
    }};

    std::vector<PerfBenchmarkResult> results;
    results.reserve(cases.size());
    bool allPassed = true;
    bool allExecuted = true;

    for (const BenchmarkCase &testCase : cases)
    {
        PerfBenchmarkResult result{};
        const std::wstring filePath = benchDir + L"\\tmp-" + std::wstring(testCase.label) + L".txt";

        if (!CreateBenchmarkFile(filePath, testCase.bytes) ||
            !RunOpenBenchmark(filePath, testCase.label, testCase.budgetMs, result))
        {
            result.label = testCase.label;
            result.bytes = testCase.bytes;
            result.budgetMs = testCase.budgetMs;
            result.success = false;
            allExecuted = false;
            allPassed = false;
        }
        else if (!result.passed)
        {
            allPassed = false;
        }

        DeleteFileW(filePath.c_str());
        results.push_back(std::move(result));
    }

    PerfBenchmarkResult typingResult{};
    if (!RunTypingBurstBenchmark(L"Typing Burst (1.2k inserts)", 450.0, typingResult))
    {
        typingResult.label = L"Typing Burst (1.2k inserts)";
        typingResult.budgetMs = 450.0;
        typingResult.success = false;
        allExecuted = false;
        allPassed = false;
    }
    else if (!typingResult.passed)
    {
        allPassed = false;
    }
    results.push_back(std::move(typingResult));

    PerfBenchmarkResult scrollResult{};
    if (!RunScrollStressBenchmark(L"Scroll Stress (3k lines)", 500.0, scrollResult))
    {
        scrollResult.label = L"Scroll Stress (3k lines)";
        scrollResult.budgetMs = 500.0;
        scrollResult.success = false;
        allExecuted = false;
        allPassed = false;
    }
    else if (!scrollResult.passed)
    {
        allPassed = false;
    }
    results.push_back(std::move(scrollResult));

    const std::wstring report = FormatBenchmarkReport(results);
    const std::wstring reportPath = benchDir + L"\\benchmark-" + BenchmarkTimestamp() + L".txt";
    const bool saved = SaveBenchmarkReport(reportPath, report);

    if (outReportPath)
        *outReportPath = reportPath;
    if (outAllPassed)
        *outAllPassed = allPassed;
    if (outAllExecuted)
        *outAllExecuted = allExecuted;

    if (!interactive)
        return saved && allExecuted && allPassed;

    const auto &lang = GetLangStrings();
    if (oldCursor)
        SetCursor(oldCursor);

    std::wstring summary;
    if (allExecuted && allPassed)
    {
        summary = lang.msgBenchmarkCompletedPass + L"\n\n";
    }
    else if (allExecuted)
    {
        summary = lang.msgBenchmarkCompletedWarn + L"\n" +
                  lang.msgBenchmarkWarnExceededBudget + L"\n\n";
    }
    else
    {
        summary = lang.msgBenchmarkCompletedFail + L"\n" +
                  lang.msgBenchmarkFailCaseNotRun + L"\n\n";
    }

    if (saved)
    {
        summary += lang.msgBenchmarkReportSavedTo;
        summary += L"\n";
        summary += reportPath;
        summary += L"\n\n";
        summary += lang.msgBenchmarkOpenReportNow;
        const int choice = MessageBoxW(g_hwndMain, summary.c_str(), lang.appName.c_str(), MB_YESNO | MB_ICONINFORMATION);
        if (choice == IDYES)
            ShellExecuteW(nullptr, L"open", reportPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
        return allExecuted && allPassed;
    }

    summary += lang.msgBenchmarkReportSaveFailed;
    MessageBoxW(g_hwndMain, summary.c_str(), lang.appName.c_str(), MB_OK | MB_ICONWARNING);
    return false;
}

void HelpRunPerformanceBenchmark()
{
    RunPerformanceBenchmark(true);
}
#endif

void ViewAlwaysOnTop()
{
    g_state.alwaysOnTop = !g_state.alwaysOnTop;
    CheckMenuItem(GetMenu(g_hwndMain), IDM_VIEW_ALWAYSONTOP, g_state.alwaysOnTop ? MF_CHECKED : MF_UNCHECKED);
    SetWindowPos(g_hwndMain, g_state.alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    SaveFontSettings();
}
