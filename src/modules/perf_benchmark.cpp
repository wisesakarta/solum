/*
  Technical Standard

  Internal performance benchmark runner used by --benchmark-ci and Help menu.
*/

#include "commands.h"

#include "core/globals.h"
#include "editor.h"
#include "file.h"
#include "settings.h"
#include "lang/lang.h"
#include "tab_layout.h"

#include <richedit.h>
#include <array>
#include <sstream>
#include <string>
#include <vector>

namespace
{
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
    SendMessageW(hwnd, WM_SETREDRAW, FALSE, 0);
    SendMessageW(hwnd, EM_SETSEL, 0, -1);
    SendMessageW(hwnd, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(L""));
    SendMessageW(hwnd, EM_SETSEL, static_cast<WPARAM>(-1), static_cast<LPARAM>(-1));

    QueryPerformanceCounter(&start);
    for (int i = 0; i < kIterations; ++i)
    {
        SendMessageW(hwnd, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(kChunk));
    }
    QueryPerformanceCounter(&end);
    SendMessageW(hwnd, WM_SETREDRAW, TRUE, 0);

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

    SendMessageW(hwnd, WM_SETREDRAW, FALSE, 0);
    SendMessageW(hwnd, EM_SETSEL, 0, 0);
    QueryPerformanceCounter(&start);
    for (int i = 0; i < kScrollOps; ++i)
        SendMessageW(hwnd, EM_LINESCROLL, 0, 1);
    QueryPerformanceCounter(&end);
    SendMessageW(hwnd, WM_SETREDRAW, TRUE, 0);

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
