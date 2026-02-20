/*
  Saka Studio & Engineering

  Menu command handlers for File, Edit, Format, and View menu operations.
  Bridges user actions to core functionality modules.
*/

#include "commands.h"
#include "core/globals.h"
#include "editor.h"
#include "file.h"
#include "ui.h"
#include "settings.h"
#include "resource.h"
#include "lang/lang.h"
#include <commdlg.h>
#include <richedit.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <vector>

namespace
{
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
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = L"Text Documents (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
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
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = L"Text Documents (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
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
    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t buf[64];
    wsprintfW(buf, L"%02d:%02d %s %02d/%02d/%04d",
              st.wHour % 12 == 0 ? 12 : st.wHour % 12, st.wMinute,
              st.wHour >= 12 ? L"PM" : L"AM", st.wMonth, st.wDay, st.wYear);
    SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(buf));
}

void FormatWordWrap()
{
    g_state.wordWrap = !g_state.wordWrap;
    CheckMenuItem(GetMenu(g_hwndMain), IDM_FORMAT_WORDWRAP, g_state.wordWrap ? MF_CHECKED : MF_UNCHECKED);
    ApplyWordWrap();
    SaveFontSettings();
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

void ViewChangeIcon()
{
    wchar_t path[MAX_PATH] = {0};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = L"Icon Files (*.ico)\0*.ico\0All Files (*.*)\0*.*\0";
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
            MessageBoxW(g_hwndMain, L"Failed to load selected icon.", lang.appName.c_str(), MB_ICONERROR);
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

void HelpCheckUpdates()
{
    ShellExecuteW(nullptr, L"open", L"https://github.com/ForLoopCodes/legacy-notepad/releases/latest", nullptr, nullptr, SW_SHOWNORMAL);
}

void ViewAlwaysOnTop()
{
    g_state.alwaysOnTop = !g_state.alwaysOnTop;
    CheckMenuItem(GetMenu(g_hwndMain), IDM_VIEW_ALWAYSONTOP, g_state.alwaysOnTop ? MF_CHECKED : MF_UNCHECKED);
    SetWindowPos(g_hwndMain, g_state.alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    SaveFontSettings();
}
