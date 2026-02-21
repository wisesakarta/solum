/*
  Saka Studio & Engineering

  Main entry point and window procedure for Saka Note text editor application.
  Coordinates all modules and handles Windows message loop and command dispatching.
*/

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <commdlg.h>
#include <richedit.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <gdiplus.h>
#include <algorithm>
#include <cstdint>
#include <vector>
#include <cwctype>

#include "resource.h"
#include "core/types.h"
#include "core/globals.h"
#include "modules/theme.h"
#include "modules/editor.h"
#include "modules/file.h"
#include "modules/ui.h"
#include "modules/background.h"
#include "modules/dialog.h"
#include "modules/commands.h"
#include "modules/settings.h"
#include "modules/menu.h"
#include "lang/lang.h"

static std::wstring MenuLabelForContext(const std::wstring &menuText)
{
    std::wstring cleaned = menuText;
    const size_t tabPos = cleaned.find(L'\t');
    if (tabPos != std::wstring::npos)
        cleaned.erase(tabPos);
    cleaned.erase(std::remove(cleaned.begin(), cleaned.end(), L'&'), cleaned.end());
    return cleaned;
}

struct DocumentTabState
{
    std::wstring text;
    std::wstring filePath;
    bool modified = false;
    Encoding encoding = Encoding::UTF8;
    LineEnding lineEnding = LineEnding::CRLF;
    bool largeFileMode = false;
    size_t sourceBytes = 0;
    bool needsReloadFromDisk = false;
};

static std::vector<DocumentTabState> g_documents;
static int g_activeDocument = -1;
static bool g_switchingDocument = false;
static std::vector<DocumentTabState> g_closedDocuments;
static constexpr size_t kMaxClosedDocuments = 20;
static bool g_updatingTabs = false;
static bool g_draggingTab = false;
static int g_dragTabIndex = -1;
static int g_hoverTabIndex = -1;
static bool g_hoverTabClose = false;
static bool g_trackingTabsMouse = false;
static bool g_tabsCustomDrawObserved = false;
static int g_tabsDpi = 96;
static HFONT g_hTabFontRegular = nullptr;
static HFONT g_hTabFontActive = nullptr;
static constexpr DWORD kSessionMagic = 0x4C4E5331; // "LNS1"
static constexpr DWORD kSessionVersion = 1;
static constexpr DWORD kSessionMaxDocuments = 64;
static constexpr DWORD kSessionMaxStringChars = 8 * 1024 * 1024;
static constexpr UINT_PTR kSessionAutosaveTimerId = 0x4C4E01;
static constexpr UINT kSessionAutosaveIntervalMs = 1500;
static constexpr DWORD kSessionRetryBackoffMs = 10000;
static bool g_sessionDirty = false;
static bool g_sessionPersisting = false;
static DWORD g_sessionRetryAtTick = 0;

static void MarkSessionDirty()
{
    g_sessionDirty = true;
    g_sessionRetryAtTick = 0;
}

static void UpdateSessionAutosaveTimer()
{
    if (!g_hwndMain)
        return;

    if (g_state.startupBehavior == StartupBehavior::ResumeAll)
        SetTimer(g_hwndMain, kSessionAutosaveTimerId, kSessionAutosaveIntervalMs, nullptr);
    else
        KillTimer(g_hwndMain, kSessionAutosaveTimerId);
}

static void UpdateRuntimeMenuStates();
static void RebuildTabsControl();
static void LoadStateFromDocument(int index);
static bool OpenPathInTabs(const std::wstring &path, bool forceReplaceCurrent = false);
static void RefreshTabsDpi();
static void ResetActiveDocumentToUntitled();
static bool ConfirmCloseForCurrentStartupBehavior();
static std::wstring ToWin32IoPath(const std::wstring &path);
static bool PathExistsForSession(const std::wstring &path);

static UINT StartupBehaviorMenuId(StartupBehavior behavior)
{
    switch (behavior)
    {
    case StartupBehavior::Classic:
        return IDM_VIEW_STARTUP_CLASSIC;
    case StartupBehavior::ResumeSaved:
        return IDM_VIEW_STARTUP_RESUMESAVED;
    case StartupBehavior::ResumeAll:
    default:
        return IDM_VIEW_STARTUP_RESUMEALL;
    }
}

static size_t EstimateTextBytes(const std::wstring &text)
{
    return text.size() * sizeof(wchar_t);
}

static bool ShouldUseLargeFileMode(size_t bytes)
{
    return bytes >= LARGE_FILE_MODE_THRESHOLD_BYTES;
}

static void EnsureSingleDocumentModel(bool captureEditorText)
{
    DocumentTabState doc;
    if (captureEditorText || g_documents.empty() || g_activeDocument < 0 || g_activeDocument >= static_cast<int>(g_documents.size()))
        doc.text = GetEditorText();
    else
        doc.text = g_documents[g_activeDocument].text;

    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    doc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateTextBytes(doc.text);
    doc.largeFileMode = g_state.largeFileMode || ShouldUseLargeFileMode(doc.sourceBytes);

    g_documents.clear();
    g_documents.push_back(std::move(doc));
    g_activeDocument = 0;

    if (g_hwndTabs)
    {
        g_updatingTabs = true;
        TabCtrl_DeleteAllItems(g_hwndTabs);
        g_updatingTabs = false;
        g_hoverTabIndex = -1;
        g_hoverTabClose = false;
    }
}

static void ApplyTabsMode(bool enabled)
{
    g_state.useTabs = enabled;

    if (!g_state.useTabs)
    {
        EnsureSingleDocumentModel(true);
    }
    else
    {
        if (g_documents.empty())
            EnsureSingleDocumentModel(true);
        RebuildTabsControl();
        if (g_activeDocument >= 0 && g_activeDocument < static_cast<int>(g_documents.size()))
            LoadStateFromDocument(g_activeDocument);
    }

    if (g_hwndTabs)
        ShowWindow(g_hwndTabs, g_state.useTabs ? SW_SHOW : SW_HIDE);

    UpdateRuntimeMenuStates();
    ResizeControls();
    UpdateStatus();
    InvalidateRect(g_hwndMain, nullptr, TRUE);
    MarkSessionDirty();
}

static void ResetActiveDocumentToUntitled()
{
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

static int ScaleTabsPx(int px)
{
    return MulDiv(px, g_tabsDpi, 96);
}

static void DestroyTabFonts()
{
    if (g_hTabFontRegular)
    {
        DeleteObject(g_hTabFontRegular);
        g_hTabFontRegular = nullptr;
    }
    if (g_hTabFontActive)
    {
        DeleteObject(g_hTabFontActive);
        g_hTabFontActive = nullptr;
    }
}

static void RefreshTabsDpi()
{
    HWND ref = g_hwndTabs ? g_hwndTabs : g_hwndMain;
    if (!ref)
    {
        g_tabsDpi = 96;
        return;
    }

    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32)
    {
        typedef UINT(WINAPI * fnGetDpiForWindow)(HWND);
#ifdef __GNUC__
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wcast-function-type"
#endif
        auto getDpiForWindow = reinterpret_cast<fnGetDpiForWindow>(GetProcAddress(hUser32, "GetDpiForWindow"));
#ifdef __GNUC__
#pragma GCC diagnostic pop
#endif
        if (getDpiForWindow)
        {
            UINT dpi = getDpiForWindow(ref);
            if (dpi >= 48 && dpi <= 960)
            {
                g_tabsDpi = static_cast<int>(dpi);
                return;
            }
        }
    }

    HDC hdc = GetDC(ref);
    if (!hdc)
    {
        g_tabsDpi = 96;
        return;
    }

    const int dpi = GetDeviceCaps(hdc, LOGPIXELSX);
    ReleaseDC(ref, hdc);
    g_tabsDpi = (dpi > 0) ? dpi : 96;
}

static void RefreshTabsVisualMetrics()
{
    if (!g_hwndTabs)
        return;

    DestroyTabFonts();

    LOGFONTW baseLf{};
    NONCLIENTMETRICSW ncm{};
    ncm.cbSize = sizeof(ncm);
    if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0))
    {
        baseLf = ncm.lfMessageFont;
    }
    else
    {
        HFONT stock = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        if (stock)
            GetObjectW(stock, sizeof(baseLf), &baseLf);
    }

    baseLf.lfHeight = -ScaleTabsPx(12);
    baseLf.lfWeight = FW_NORMAL;
    baseLf.lfQuality = CLEARTYPE_QUALITY;
    g_hTabFontRegular = CreateFontIndirectW(&baseLf);

    LOGFONTW activeLf = baseLf;
    activeLf.lfWeight = FW_SEMIBOLD;
    g_hTabFontActive = CreateFontIndirectW(&activeLf);

    HFONT effectiveFont = g_hTabFontRegular ? g_hTabFontRegular : reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    SendMessageW(g_hwndTabs, WM_SETFONT, reinterpret_cast<WPARAM>(effectiveFont), TRUE);
    SendMessageW(g_hwndTabs, TCM_SETPADDING, 0, MAKELPARAM(ScaleTabsPx(12), ScaleTabsPx(4)));
}

static LRESULT CALLBACK MenuPopupCbtHookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HCBT_CREATEWND)
    {
        HWND hwnd = reinterpret_cast<HWND>(wParam);
        wchar_t className[32] = {};
        if (GetClassNameW(hwnd, className, static_cast<int>(std::size(className))) > 0 &&
            wcscmp(className, L"#32768") == 0)
        {
            // Try to disable DWM-heavy effects on popup menu windows.
            HMODULE hDwmapi = LoadLibraryW(L"dwmapi.dll");
            if (hDwmapi)
            {
                typedef HRESULT(WINAPI * fnDwmSetWindowAttribute)(HWND, DWORD, LPCVOID, DWORD);
#ifdef __GNUC__
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wcast-function-type"
#endif
                auto dwmSetWindowAttribute = reinterpret_cast<fnDwmSetWindowAttribute>(GetProcAddress(hDwmapi, "DwmSetWindowAttribute"));
#ifdef __GNUC__
#pragma GCC diagnostic pop
#endif
                if (dwmSetWindowAttribute)
                {
                    const DWMNCRENDERINGPOLICY noNcRendering = DWMNCRP_DISABLED;
                    dwmSetWindowAttribute(hwnd, DWMWA_NCRENDERING_POLICY, &noNcRendering, sizeof(noNcRendering));
                }
                FreeLibrary(hDwmapi);
            }

            const LONG_PTR classStyle = GetClassLongPtrW(hwnd, GCL_STYLE);
            if (classStyle & CS_DROPSHADOW)
                SetClassLongPtrW(hwnd, GCL_STYLE, classStyle & ~static_cast<LONG_PTR>(CS_DROPSHADOW));
        }
    }
    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

static UINT TrackPopupMenuLightweight(HMENU hPopup, UINT flags, int x, int y, HWND hwndOwner)
{
    if (!hPopup || !hwndOwner)
        return 0;

    HHOOK hook = SetWindowsHookExW(WH_CBT, MenuPopupCbtHookProc, nullptr, GetCurrentThreadId());
    const UINT cmd = TrackPopupMenu(hPopup,
                                    flags | TPM_NOANIMATION,
                                    x, y, 0, hwndOwner, nullptr);
    if (hook)
        UnhookWindowsHookEx(hook);
    return cmd;
}

static void DrawCloseGlyph(HDC hdc, const RECT &rc, COLORREF color);

static bool WriteAllBytes(HANDLE hFile, const void *data, DWORD bytes)
{
    const BYTE *cursor = reinterpret_cast<const BYTE *>(data);
    DWORD remaining = bytes;
    while (remaining > 0)
    {
        DWORD written = 0;
        if (!WriteFile(hFile, cursor, remaining, &written, nullptr))
            return false;
        if (written == 0)
            return false;
        cursor += written;
        remaining -= written;
    }
    return true;
}

static bool ReadAllBytes(HANDLE hFile, void *data, DWORD bytes)
{
    BYTE *cursor = reinterpret_cast<BYTE *>(data);
    DWORD remaining = bytes;
    while (remaining > 0)
    {
        DWORD read = 0;
        if (!ReadFile(hFile, cursor, remaining, &read, nullptr))
            return false;
        if (read == 0)
            return false;
        cursor += read;
        remaining -= read;
    }
    return true;
}

static bool WriteUInt32(HANDLE hFile, DWORD value)
{
    return WriteAllBytes(hFile, &value, sizeof(value));
}

static bool ReadUInt32(HANDLE hFile, DWORD &value)
{
    return ReadAllBytes(hFile, &value, sizeof(value));
}

static bool WriteWideString(HANDLE hFile, const std::wstring &value)
{
    if (value.size() > kSessionMaxStringChars)
        return false;

    const DWORD charCount = static_cast<DWORD>(value.size());
    if (!WriteUInt32(hFile, charCount))
        return false;
    if (charCount == 0)
        return true;
    return WriteAllBytes(hFile, value.data(), charCount * sizeof(wchar_t));
}

static bool ReadWideString(HANDLE hFile, std::wstring &value)
{
    DWORD charCount = 0;
    if (!ReadUInt32(hFile, charCount))
        return false;
    if (charCount > kSessionMaxStringChars)
        return false;

    value.clear();
    if (charCount == 0)
        return true;

    value.resize(charCount);
    return ReadAllBytes(hFile, value.data(), charCount * sizeof(wchar_t));
}

static bool WriteDocumentRecord(HANDLE hFile, const DocumentTabState &doc)
{
    const DWORD modifiedFlag = doc.modified ? 1u : 0u;
    if (!WriteUInt32(hFile, modifiedFlag))
        return false;
    if (!WriteUInt32(hFile, static_cast<DWORD>(doc.encoding)))
        return false;
    if (!WriteUInt32(hFile, static_cast<DWORD>(doc.lineEnding)))
        return false;
    if (!WriteWideString(hFile, doc.filePath))
        return false;
    const bool persistText = doc.filePath.empty() || doc.modified;
    const std::wstring &textToPersist = persistText ? doc.text : std::wstring();
    if (!WriteWideString(hFile, textToPersist))
        return false;
    return true;
}

static bool ReadDocumentRecord(HANDLE hFile, DocumentTabState &doc)
{
    DWORD modifiedFlag = 0;
    DWORD encodingValue = 0;
    DWORD lineEndingValue = 0;
    if (!ReadUInt32(hFile, modifiedFlag))
        return false;
    if (!ReadUInt32(hFile, encodingValue))
        return false;
    if (!ReadUInt32(hFile, lineEndingValue))
        return false;
    if (!ReadWideString(hFile, doc.filePath))
        return false;
    if (!ReadWideString(hFile, doc.text))
        return false;

    doc.modified = (modifiedFlag != 0);
    if (encodingValue <= static_cast<DWORD>(Encoding::ANSI))
        doc.encoding = static_cast<Encoding>(encodingValue);
    else
        doc.encoding = Encoding::UTF8;

    if (lineEndingValue <= static_cast<DWORD>(LineEnding::CR))
        doc.lineEnding = static_cast<LineEnding>(lineEndingValue);
    else
        doc.lineEnding = LineEnding::CRLF;

    doc.needsReloadFromDisk = false;
    if (!doc.modified && doc.text.empty() && !doc.filePath.empty())
    {
        doc.sourceBytes = 0;
        doc.largeFileMode = false;
        doc.needsReloadFromDisk = true;
        return true;
    }

    doc.sourceBytes = EstimateTextBytes(doc.text);
    doc.largeFileMode = ShouldUseLargeFileMode(doc.sourceBytes);
    return true;
}

static bool LoadDocumentTextFromDisk(DocumentTabState &doc)
{
    if (doc.filePath.empty())
        return false;

    const std::wstring ioPath = ToWin32IoPath(doc.filePath);
    HANDLE hFile = CreateFileW(ioPath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
                               OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
        return false;

    LARGE_INTEGER fileSize = {};
    if (!GetFileSizeEx(hFile, &fileSize) || fileSize.QuadPart < 0 || fileSize.QuadPart > static_cast<LONGLONG>(MAXDWORD))
    {
        CloseHandle(hFile);
        return false;
    }

    const DWORD size = static_cast<DWORD>(fileSize.QuadPart);
    std::vector<BYTE> data(size);
    if (size > 0 && !ReadAllBytes(hFile, data.data(), size))
    {
        CloseHandle(hFile);
        return false;
    }
    CloseHandle(hFile);

    auto [enc, le] = DetectEncoding(data);
    doc.text = DecodeText(data, enc);
    doc.encoding = enc;
    doc.lineEnding = le;
    doc.sourceBytes = static_cast<size_t>(size);
    doc.largeFileMode = ShouldUseLargeFileMode(doc.sourceBytes);
    doc.needsReloadFromDisk = false;
    doc.modified = false;
    return true;
}

static std::wstring SessionFilePath()
{
    wchar_t localAppData[MAX_PATH] = {};
    DWORD len = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);

    std::wstring dirPath;
    if (len > 0 && len < MAX_PATH)
    {
        dirPath = localAppData;
    }
    else
    {
        wchar_t modulePath[MAX_PATH] = {};
        if (GetModuleFileNameW(nullptr, modulePath, MAX_PATH) == 0)
            return L"session.dat";
        PathRemoveFileSpecW(modulePath);
        dirPath = modulePath;
    }

    dirPath += L"\\SakaNote";
    CreateDirectoryW(dirPath.c_str(), nullptr);
    return dirPath + L"\\session.dat";
}

static std::wstring ToWin32IoPath(const std::wstring &path)
{
    if (path.empty())
        return path;
    if (path.rfind(L"\\\\?\\", 0) == 0)
        return path;

    if (path.rfind(L"\\\\", 0) == 0)
        return L"\\\\?\\UNC\\" + path.substr(2);

    if (path.size() >= MAX_PATH && path.size() > 2 && path[1] == L':' &&
        (path[2] == L'\\' || path[2] == L'/'))
        return L"\\\\?\\" + path;

    return path;
}

static bool PathExistsForSession(const std::wstring &path)
{
    if (path.empty())
        return false;

    const std::wstring ioPath = ToWin32IoPath(path);
    const DWORD attrs = GetFileAttributesW(ioPath.c_str());
    return attrs != INVALID_FILE_ATTRIBUTES && ((attrs & FILE_ATTRIBUTE_DIRECTORY) == 0);
}

static std::wstring NormalizePathForCompare(const std::wstring &path)
{
    if (path.empty())
        return {};

    wchar_t fullPath[MAX_PATH] = {};
    DWORD len = GetFullPathNameW(path.c_str(), MAX_PATH, fullPath, nullptr);
    std::wstring normalized;
    if (len > 0 && len < MAX_PATH)
        normalized = fullPath;
    else
        normalized = path;

    std::transform(normalized.begin(), normalized.end(), normalized.begin(), towlower);
    return normalized;
}

static std::wstring DocumentTabLabel(const DocumentTabState &doc)
{
    const auto &lang = GetLangStrings();
    std::wstring label = doc.filePath.empty() ? lang.untitled : PathFindFileNameW(doc.filePath.c_str());
    if (doc.modified)
        label.insert(label.begin(), L'*');
    return label;
}

static RECT TabCloseRect(const RECT &itemRect)
{
    RECT rc = itemRect;
    const int closeSize = ScaleTabsPx(14);
    const int rightPadding = ScaleTabsPx(8);
    const int topOffset = ((rc.bottom - rc.top) - closeSize) / 2;
    rc.right -= rightPadding;
    rc.left = rc.right - closeSize;
    rc.top += topOffset;
    rc.bottom = rc.top + closeSize;
    return rc;
}

static RECT TabTextRect(const RECT &itemRect)
{
    RECT rc = itemRect;
    const RECT closeRc = TabCloseRect(itemRect);
    rc.left += ScaleTabsPx(10);
    rc.right = closeRc.left - ScaleTabsPx(8);
    if (rc.right < rc.left)
        rc.right = rc.left;
    return rc;
}

static void SetDocumentTabLabel(int index)
{
    if (!g_hwndTabs || index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    TCITEMW item{};
    item.mask = TCIF_TEXT;
    std::wstring label = DocumentTabLabel(g_documents[index]);
    item.pszText = label.data();
    TabCtrl_SetItem(g_hwndTabs, index, &item);
}

static void RebuildTabsControl()
{
    if (!g_hwndTabs)
        return;

    g_hoverTabIndex = -1;
    g_hoverTabClose = false;
    g_updatingTabs = true;
    TabCtrl_DeleteAllItems(g_hwndTabs);

    for (int i = 0; i < static_cast<int>(g_documents.size()); ++i)
    {
        TCITEMW item{};
        item.mask = TCIF_TEXT;
        std::wstring label = DocumentTabLabel(g_documents[i]);
        item.pszText = label.data();
        TabCtrl_InsertItem(g_hwndTabs, i, &item);
    }

    if (g_activeDocument >= 0 && g_activeDocument < static_cast<int>(g_documents.size()))
        TabCtrl_SetCurSel(g_hwndTabs, g_activeDocument);
    g_updatingTabs = false;
}

static bool IsTabCloseRectHit(int index, POINT ptClient)
{
    if (!g_hwndTabs || index < 0 || index >= TabCtrl_GetItemCount(g_hwndTabs))
        return false;

    RECT rc{};
    if (!TabCtrl_GetItemRect(g_hwndTabs, index, &rc))
        return false;

    RECT closeRc = TabCloseRect(rc);
    return PtInRect(&closeRc, ptClient) != 0;
}

static bool IsTabCloseHotspot(int index, POINT ptClient)
{
    // Match modern notepad behavior: close affordance is active only while hovered.
    if (index != g_hoverTabIndex)
        return false;
    return IsTabCloseRectHit(index, ptClient);
}

static void SyncDocumentFromState(int index, bool includeText)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    DocumentTabState &doc = g_documents[index];
    if (includeText)
    {
        doc.text = GetEditorText();
        doc.needsReloadFromDisk = false;
    }
    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    doc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateTextBytes(doc.text);
    doc.largeFileMode = g_state.largeFileMode || ShouldUseLargeFileMode(doc.sourceBytes);
    SetDocumentTabLabel(index);
}

static void LoadStateFromDocument(int index)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    g_switchingDocument = true;
    DocumentTabState &doc = g_documents[index];
    if (doc.needsReloadFromDisk)
    {
        if (!LoadDocumentTextFromDisk(doc))
        {
            doc.needsReloadFromDisk = false;
            doc.text.clear();
            doc.sourceBytes = 0;
            doc.largeFileMode = false;
        }
    }

    const bool wrapModeChanged = (g_state.largeFileMode != doc.largeFileMode);
    g_state.filePath = doc.filePath;
    g_state.modified = doc.modified;
    g_state.encoding = doc.encoding;
    g_state.lineEnding = doc.lineEnding;
    g_state.largeFileMode = doc.largeFileMode;
    g_state.largeFileBytes = doc.sourceBytes;
    if (wrapModeChanged)
        ApplyWordWrap();
    SetEditorText(doc.text);
    UpdateTitle();
    UpdateStatus();
    SetDocumentTabLabel(index);
    g_switchingDocument = false;
    SetFocus(g_hwndEditor);
}

static void SwitchToDocument(int index)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;
    if (index == g_activeDocument)
        return;

    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    g_activeDocument = index;
    if (g_hwndTabs)
        TabCtrl_SetCurSel(g_hwndTabs, index);
    LoadStateFromDocument(index);
    MarkSessionDirty();
}

static void RefreshAllDocumentTabLabels()
{
    for (int i = 0; i < static_cast<int>(g_documents.size()); ++i)
        SetDocumentTabLabel(i);
}

static void CreateInitialDocumentTabIfNeeded()
{
    if (!g_hwndTabs || !g_documents.empty())
        return;

    DocumentTabState doc;
    doc.text = GetEditorText();
    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    doc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateTextBytes(doc.text);
    doc.largeFileMode = g_state.largeFileMode || ShouldUseLargeFileMode(doc.sourceBytes);

    g_documents.push_back(doc);
    g_activeDocument = 0;

    TCITEMW item{};
    item.mask = TCIF_TEXT;
    std::wstring label = DocumentTabLabel(doc);
    item.pszText = label.data();
    TabCtrl_InsertItem(g_hwndTabs, 0, &item);
    TabCtrl_SetCurSel(g_hwndTabs, 0);
    UpdateRuntimeMenuStates();
}

static void CreateNewDocumentTab()
{
    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    DocumentTabState doc;
    g_documents.push_back(doc);
    int index = static_cast<int>(g_documents.size()) - 1;

    if (g_hwndTabs)
    {
        TCITEMW item{};
        item.mask = TCIF_TEXT;
        std::wstring label = DocumentTabLabel(doc);
        item.pszText = label.data();
        TabCtrl_InsertItem(g_hwndTabs, index, &item);
        TabCtrl_SetCurSel(g_hwndTabs, index);
    }

    g_activeDocument = index;
    LoadStateFromDocument(index);
    UpdateRuntimeMenuStates();
    MarkSessionDirty();
}

static int FindDocumentByPath(const std::wstring &path)
{
    const std::wstring needle = NormalizePathForCompare(path);
    if (needle.empty())
        return -1;

    for (int i = 0; i < static_cast<int>(g_documents.size()); ++i)
    {
        if (NormalizePathForCompare(g_documents[i].filePath) == needle)
            return i;
    }
    return -1;
}

static bool IsCurrentDocumentEmptyAndUntitled()
{
    if (g_activeDocument < 0 || g_activeDocument >= static_cast<int>(g_documents.size()))
        return false;

    const DocumentTabState &doc = g_documents[g_activeDocument];
    return doc.filePath.empty() && !doc.modified && doc.text.empty();
}

static bool OpenPathInTabs(const std::wstring &path, bool forceReplaceCurrent)
{
    if (path.empty())
        return false;

    if (!g_state.useTabs)
    {
        if (!forceReplaceCurrent && !ConfirmDiscard())
            return false;
        if (!LoadFile(path))
            return false;
        EnsureSingleDocumentModel(true);
        UpdateRuntimeMenuStates();
        MarkSessionDirty();
        return true;
    }

    const int existingIndex = FindDocumentByPath(path);
    if (existingIndex >= 0)
    {
        SwitchToDocument(existingIndex);
        return true;
    }

    const int previousIndex = g_activeDocument;
    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    bool createdNewTab = false;
    if (!IsCurrentDocumentEmptyAndUntitled())
    {
        CreateNewDocumentTab();
        createdNewTab = true;
    }

    if (!LoadFile(path))
    {
        if (createdNewTab && !g_documents.empty())
        {
            const int failedIndex = g_activeDocument;
            if (failedIndex >= 0 && failedIndex < static_cast<int>(g_documents.size()))
            {
                g_documents.erase(g_documents.begin() + failedIndex);
                if (g_hwndTabs)
                {
                    g_updatingTabs = true;
                    TabCtrl_DeleteItem(g_hwndTabs, failedIndex);
                    g_updatingTabs = false;
                }
            }

            if (g_documents.empty())
            {
                DocumentTabState fallbackDoc;
                g_documents.push_back(std::move(fallbackDoc));
                g_activeDocument = 0;
                RebuildTabsControl();
                LoadStateFromDocument(0);
            }
            else
            {
                int restoreIndex = previousIndex;
                if (restoreIndex < 0 || restoreIndex >= static_cast<int>(g_documents.size()))
                    restoreIndex = 0;
                g_activeDocument = restoreIndex;
                if (g_hwndTabs)
                    TabCtrl_SetCurSel(g_hwndTabs, restoreIndex);
                LoadStateFromDocument(restoreIndex);
            }

            UpdateRuntimeMenuStates();
        }
        return false;
    }

    SyncDocumentFromState(g_activeDocument, true);
    UpdateRuntimeMenuStates();
    MarkSessionDirty();
    return true;
}

static bool SaveOpenDocumentSession(bool persistPathFallback)
{
    if (g_sessionPersisting)
        return false;
    g_sessionPersisting = true;

    bool saveOk = true;

    if (g_activeDocument >= 0 && g_activeDocument < static_cast<int>(g_documents.size()))
    {
        DocumentTabState &activeDoc = g_documents[g_activeDocument];
        if (g_state.modified || g_state.filePath.empty())
        {
            activeDoc.text = GetEditorText();
            activeDoc.needsReloadFromDisk = false;
        }
        else
        {
            activeDoc.needsReloadFromDisk = false;
        }
        activeDoc.filePath = g_state.filePath;
        activeDoc.modified = g_state.modified;
        activeDoc.encoding = g_state.encoding;
        activeDoc.lineEnding = g_state.lineEnding;
        activeDoc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateTextBytes(activeDoc.text);
        activeDoc.largeFileMode = g_state.largeFileMode || ShouldUseLargeFileMode(activeDoc.sourceBytes);
    }

    const bool resumeAll = (g_state.startupBehavior == StartupBehavior::ResumeAll);
    const bool startupClassic = (g_state.startupBehavior == StartupBehavior::Classic);
    const std::wstring sessionFilePath = SessionFilePath();

    if (startupClassic)
    {
        if (persistPathFallback)
        {
            DeleteFileW(sessionFilePath.c_str());
            SaveOpenTabsSession({}, -1);
        }
        g_sessionPersisting = false;
        return true;
    }

    if (resumeAll)
    {
        HANDLE hFile = CreateFileW(sessionFilePath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            const DWORD docCount = static_cast<DWORD>(std::min<size_t>(g_documents.size(), kSessionMaxDocuments));
            DWORD activeDocIndex = 0xFFFFFFFFu;
            if (g_activeDocument >= 0 && g_activeDocument < static_cast<int>(docCount))
                activeDocIndex = static_cast<DWORD>(g_activeDocument);

            bool ok = WriteUInt32(hFile, kSessionMagic) &&
                      WriteUInt32(hFile, kSessionVersion) &&
                      WriteUInt32(hFile, docCount) &&
                      WriteUInt32(hFile, activeDocIndex);
            for (DWORD i = 0; ok && i < docCount; ++i)
                ok = WriteDocumentRecord(hFile, g_documents[i]);

            CloseHandle(hFile);
            if (!ok)
            {
                DeleteFileW(sessionFilePath.c_str());
                saveOk = false;
            }
        }
        else
        {
            saveOk = false;
        }
    }
    else
    {
        DeleteFileW(sessionFilePath.c_str());
    }

    if (persistPathFallback)
    {
        std::vector<std::wstring> sessionPaths;
        std::vector<std::wstring> normalizedPaths;
        sessionPaths.reserve(g_documents.size());
        normalizedPaths.reserve(g_documents.size());

        int activePathIndex = -1;
        for (int i = 0; i < static_cast<int>(g_documents.size()); ++i)
        {
            const std::wstring &path = g_documents[i].filePath;
            if (path.empty())
                continue;

            const std::wstring normalized = NormalizePathForCompare(path);
            if (normalized.empty())
                continue;

            int existingIndex = -1;
            for (int j = 0; j < static_cast<int>(normalizedPaths.size()); ++j)
            {
                if (normalizedPaths[j] == normalized)
                {
                    existingIndex = j;
                    break;
                }
            }

            if (existingIndex >= 0)
            {
                if (i == g_activeDocument)
                    activePathIndex = existingIndex;
                continue;
            }

            normalizedPaths.push_back(normalized);
            sessionPaths.push_back(path);
            if (i == g_activeDocument)
                activePathIndex = static_cast<int>(sessionPaths.size()) - 1;
        }

        SaveOpenTabsSession(sessionPaths, activePathIndex);
    }

    g_sessionPersisting = false;
    return saveOk;
}

static bool RestoreOpenDocumentSession()
{
    if (g_state.startupBehavior == StartupBehavior::Classic)
        return false;

    const bool allowUnsavedRestore = (g_state.startupBehavior == StartupBehavior::ResumeAll);

    if (allowUnsavedRestore)
    {
    const std::wstring sessionFilePath = SessionFilePath();
    HANDLE hFile = CreateFileW(sessionFilePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile != INVALID_HANDLE_VALUE)
    {
        DWORD magic = 0;
        DWORD version = 0;
        DWORD docCount = 0;
        DWORD activeDocIndex = 0xFFFFFFFFu;
        bool ok = ReadUInt32(hFile, magic) &&
                  ReadUInt32(hFile, version) &&
                  ReadUInt32(hFile, docCount) &&
                  ReadUInt32(hFile, activeDocIndex) &&
                  magic == kSessionMagic &&
                  version == kSessionVersion &&
                  docCount > 0 &&
                  docCount <= kSessionMaxDocuments;

        std::vector<DocumentTabState> restoredDocs;
        restoredDocs.reserve(docCount);
        for (DWORD i = 0; ok && i < docCount; ++i)
        {
            DocumentTabState doc;
            ok = ReadDocumentRecord(hFile, doc);
            if (ok)
                restoredDocs.push_back(std::move(doc));
        }
        CloseHandle(hFile);

        if (ok && !restoredDocs.empty())
        {
            g_documents = std::move(restoredDocs);
            g_closedDocuments.clear();
            g_activeDocument = (activeDocIndex < g_documents.size()) ? static_cast<int>(activeDocIndex) : 0;
            if (!g_state.useTabs && g_documents.size() > 1)
            {
                DocumentTabState activeDoc = g_documents[g_activeDocument];
                g_documents.clear();
                g_documents.push_back(std::move(activeDoc));
                g_activeDocument = 0;
            }
            RebuildTabsControl();
            LoadStateFromDocument(g_activeDocument);
            RefreshAllDocumentTabLabels();
            UpdateRuntimeMenuStates();
            return true;
        }
    }
    }

    std::vector<std::wstring> sessionPaths;
    int activePathIndex = -1;
    LoadOpenTabsSession(sessionPaths, activePathIndex);
    if (sessionPaths.empty())
        return false;

    if (!g_state.useTabs)
    {
        int preferred = activePathIndex;
        if (preferred < 0 || preferred >= static_cast<int>(sessionPaths.size()))
            preferred = static_cast<int>(sessionPaths.size()) - 1;
        const std::wstring &path = sessionPaths[preferred];
        if (path.empty() || !PathExistsForSession(path))
            return false;
        return OpenPathInTabs(path);
    }

    bool openedAny = false;
    for (const auto &path : sessionPaths)
    {
        if (path.empty() || !PathExistsForSession(path))
            continue;
        if (OpenPathInTabs(path))
            openedAny = true;
    }

    if (!openedAny)
        return false;

    if (activePathIndex < 0 || activePathIndex >= static_cast<int>(sessionPaths.size()))
        return true;

    const int activeDocIndex = FindDocumentByPath(sessionPaths[activePathIndex]);
    if (activeDocIndex >= 0)
        SwitchToDocument(activeDocIndex);
    UpdateRuntimeMenuStates();
    return true;
}

static void OpenFileInNewDocumentTabDialog()
{
    wchar_t path[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = L"Text Documents (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_ENABLESIZING;
    if (GetOpenFileNameW(&ofn))
        OpenPathInTabs(path);
}

static void PushClosedDocument(const DocumentTabState &doc)
{
    if (doc.filePath.empty() && doc.text.empty() && !doc.modified)
        return;

    g_closedDocuments.push_back(doc);
    if (g_closedDocuments.size() > kMaxClosedDocuments)
        g_closedDocuments.erase(g_closedDocuments.begin());
}

static void CloseDocumentTabAt(int index)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    SwitchToDocument(index);

    if (g_documents.size() <= 1)
    {
        if (!ConfirmDiscard())
            return;
        PushClosedDocument(g_documents[g_activeDocument]);
        ResetActiveDocumentToUntitled();
        SyncDocumentFromState(g_activeDocument, true);
        UpdateRuntimeMenuStates();
        MarkSessionDirty();
        return;
    }

    if (!ConfirmDiscard())
        return;

    const int closingIndex = index;
    PushClosedDocument(g_documents[closingIndex]);
    g_documents.erase(g_documents.begin() + closingIndex);

    int nextIndex = closingIndex;
    if (nextIndex >= static_cast<int>(g_documents.size()))
        nextIndex = static_cast<int>(g_documents.size()) - 1;

    g_activeDocument = nextIndex;
    RebuildTabsControl();
    LoadStateFromDocument(nextIndex);
    RefreshAllDocumentTabLabels();
    UpdateRuntimeMenuStates();
    MarkSessionDirty();
}

static void CloseCurrentDocumentTab()
{
    CloseDocumentTabAt(g_activeDocument);
}

static bool ConfirmCloseForCurrentStartupBehavior()
{
    // Match startup behavior semantics:
    // - ResumeAll: unsaved buffers are persisted in session, no save prompt on close.
    // - Classic / ResumeSaved: unsaved changes would be lost, require confirmation.
    if (g_state.startupBehavior == StartupBehavior::ResumeAll)
        return true;

    if (g_activeDocument >= 0 && g_activeDocument < static_cast<int>(g_documents.size()))
        SyncDocumentFromState(g_activeDocument, true);

    if (!g_state.useTabs || g_documents.empty())
        return ConfirmDiscard();

    const int originalIndex = g_activeDocument;
    const int count = static_cast<int>(g_documents.size());
    for (int i = 0; i < count; ++i)
    {
        const DocumentTabState &doc = g_documents[i];
        if (!doc.modified)
            continue;
        if (doc.filePath.empty() && doc.text.empty())
            continue;

        SwitchToDocument(i);
        if (!ConfirmDiscard())
        {
            if (originalIndex >= 0 && originalIndex < static_cast<int>(g_documents.size()))
                SwitchToDocument(originalIndex);
            return false;
        }
        SyncDocumentFromState(i, true);
    }

    if (originalIndex >= 0 && originalIndex < static_cast<int>(g_documents.size()))
        SwitchToDocument(originalIndex);
    return true;
}

static void ReopenClosedDocumentTab()
{
    if (g_closedDocuments.empty())
        return;

    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    DocumentTabState doc = g_closedDocuments.back();
    g_closedDocuments.pop_back();
    g_documents.push_back(doc);
    g_activeDocument = static_cast<int>(g_documents.size()) - 1;
    RebuildTabsControl();
    LoadStateFromDocument(g_activeDocument);
    UpdateRuntimeMenuStates();
    MarkSessionDirty();
}

static void ReorderDocumentTab(int fromIndex, int toIndex)
{
    if (fromIndex < 0 || toIndex < 0 || fromIndex >= static_cast<int>(g_documents.size()) || toIndex >= static_cast<int>(g_documents.size()))
        return;
    if (fromIndex == toIndex)
        return;

    std::swap(g_documents[fromIndex], g_documents[toIndex]);
    if (g_activeDocument == fromIndex)
        g_activeDocument = toIndex;
    else if (g_activeDocument == toIndex)
        g_activeDocument = fromIndex;

    RebuildTabsControl();
    UpdateRuntimeMenuStates();
    MarkSessionDirty();
}

static void SwitchToNextDocumentTab(bool backward)
{
    if (g_documents.size() <= 1)
        return;

    int current = g_activeDocument;
    if (current < 0)
        current = 0;

    const int count = static_cast<int>(g_documents.size());
    int next = backward ? (current - 1 + count) % count : (current + 1) % count;
    SwitchToDocument(next);
}

static void UpdateRuntimeMenuStates()
{
    HMENU hMenu = GetMenu(g_hwndMain);
    if (!hMenu)
        return;

    CheckMenuItem(hMenu, IDM_VIEW_USETABS, MF_BYCOMMAND | (g_state.useTabs ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuRadioItem(hMenu,
                       IDM_VIEW_STARTUP_CLASSIC,
                       IDM_VIEW_STARTUP_RESUMESAVED,
                       StartupBehaviorMenuId(g_state.startupBehavior),
                       MF_BYCOMMAND);
    CheckMenuItem(hMenu, IDM_FORMAT_WORDWRAP, MF_BYCOMMAND | ((g_state.wordWrap && !g_state.largeFileMode) ? MF_CHECKED : MF_UNCHECKED));
    EnableMenuItem(hMenu, IDM_FORMAT_WORDWRAP, MF_BYCOMMAND | (g_state.largeFileMode ? MF_GRAYED : MF_ENABLED));

    const bool tabModeActive = g_state.useTabs;
    const bool hasMultipleTabs = g_documents.size() > 1;
    const bool hasClosedTabs = !g_closedDocuments.empty();

    EnableMenuItem(hMenu, IDM_FILE_CLOSETAB, MF_BYCOMMAND | ((tabModeActive && hasMultipleTabs) ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(hMenu, IDM_FILE_REOPENCLOSEDTAB, MF_BYCOMMAND | ((tabModeActive && hasClosedTabs) ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(hMenu, IDM_FILE_NEXTTAB, MF_BYCOMMAND | ((tabModeActive && hasMultipleTabs) ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(hMenu, IDM_FILE_PREVTAB, MF_BYCOMMAND | ((tabModeActive && hasMultipleTabs) ? MF_ENABLED : MF_GRAYED));

    DrawMenuBar(g_hwndMain);
}

static void DrawCloseGlyph(HDC hdc, const RECT &rc, COLORREF color)
{
    const int thickness = std::max(1, ScaleTabsPx(1));
    HPEN hPen = CreatePen(PS_SOLID, thickness, color);
    if (!hPen)
        return;
    HGDIOBJ oldPen = SelectObject(hdc, hPen);

    const int margin = ScaleTabsPx(4);
    MoveToEx(hdc, rc.left + margin, rc.top + margin, nullptr);
    LineTo(hdc, rc.right - margin, rc.bottom - margin);
    MoveToEx(hdc, rc.right - margin, rc.top + margin, nullptr);
    LineTo(hdc, rc.left + margin, rc.bottom - margin);

    SelectObject(hdc, oldPen);
    DeleteObject(hPen);
}

struct TabPaintPalette
{
    COLORREF stripBg;
    COLORREF stripBorder;
    COLORREF activeBg;
    COLORREF inactiveBg;
    COLORREF hoverBg;
    COLORREF borderColor;
    COLORREF textColor;
    COLORREF activeTextColor;
    COLORREF closeColor;
    COLORREF closeHoverBg;
    COLORREF closeHoverFg;
};

static TabPaintPalette GetTabPaintPalette(bool dark)
{
    if (dark)
    {
        return {
            RGB(24, 27, 32),  // stripBg
            RGB(56, 62, 72),  // stripBorder
            RGB(44, 50, 58),  // activeBg
            RGB(24, 27, 32),  // inactiveBg
            RGB(39, 44, 51),  // hoverBg
            RGB(68, 74, 84),  // borderColor
            RGB(212, 218, 226), // textColor
            RGB(247, 249, 252), // activeTextColor
            RGB(204, 210, 220), // closeColor
            RGB(170, 63, 63),   // closeHoverBg
            RGB(255, 255, 255)  // closeHoverFg
        };
    }

    return {
        RGB(246, 248, 252),  // stripBg
        RGB(214, 220, 229),  // stripBorder
        RGB(255, 255, 255),  // activeBg
        RGB(246, 248, 252),  // inactiveBg
        RGB(230, 235, 243),  // hoverBg
        RGB(203, 210, 220),  // borderColor
        RGB(70, 76, 84),     // textColor
        RGB(23, 27, 32),     // activeTextColor
        RGB(92, 97, 104),    // closeColor
        RGB(225, 79, 79),    // closeHoverBg
        RGB(255, 255, 255)   // closeHoverFg
    };
}

static void FillSolidRectDc(HDC hdc, const RECT &rc, COLORREF color)
{
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(DC_BRUSH));
    const COLORREF oldBrushColor = SetDCBrushColor(hdc, color);
    FillRect(hdc, &rc, reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
    SetDCBrushColor(hdc, oldBrushColor);
    SelectObject(hdc, oldBrush);
}

static void DrawRoundedRectDc(HDC hdc, const RECT &rc, int radius, COLORREF fillColor, COLORREF borderColor)
{
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(DC_BRUSH));
    HGDIOBJ oldPen = SelectObject(hdc, GetStockObject(DC_PEN));
    const COLORREF oldBrushColor = SetDCBrushColor(hdc, fillColor);
    const COLORREF oldPenColor = SetDCPenColor(hdc, borderColor);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);
    SetDCBrushColor(hdc, oldBrushColor);
    SetDCPenColor(hdc, oldPenColor);
    SelectObject(hdc, oldPen);
    SelectObject(hdc, oldBrush);
}

static void DrawTabStripBackground(HDC hdc, const RECT &rcClient, const TabPaintPalette &palette)
{
    FillSolidRectDc(hdc, rcClient, palette.stripBg);
    RECT separator = rcClient;
    separator.top = std::max(separator.top, separator.bottom - std::max(1, ScaleTabsPx(1)));
    RECT activeRect{};
    const bool hasActive = (g_hwndTabs && g_activeDocument >= 0 && TabCtrl_GetItemRect(g_hwndTabs, g_activeDocument, &activeRect));
    if (!hasActive)
    {
        FillSolidRectDc(hdc, separator, palette.stripBorder);
        return;
    }

    activeRect.left += ScaleTabsPx(2);
    activeRect.right -= ScaleTabsPx(2);

    RECT leftLine = separator;
    leftLine.right = std::max(leftLine.left, std::min(leftLine.right, activeRect.left));
    if (leftLine.right > leftLine.left)
        FillSolidRectDc(hdc, leftLine, palette.stripBorder);

    RECT rightLine = separator;
    rightLine.left = std::min(rightLine.right, std::max(rightLine.left, activeRect.right));
    if (rightLine.right > rightLine.left)
        FillSolidRectDc(hdc, rightLine, palette.stripBorder);
}

static void DrawTabItemVisual(HDC hdc, int index, const RECT &rawItemRect, const TabPaintPalette &palette)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    const bool selected = (index == g_activeDocument);
    const bool hovered = (index == g_hoverTabIndex);
    const bool showClose = hovered;
    const bool closeHovered = hovered && g_hoverTabClose;

    RECT contentRect = rawItemRect;
    contentRect.left += ScaleTabsPx(2);
    contentRect.right -= ScaleTabsPx(2);
    contentRect.top += ScaleTabsPx(2);
    contentRect.bottom -= selected ? 0 : ScaleTabsPx(4);

    if (selected && g_hwndTabs)
    {
        RECT strip{};
        GetClientRect(g_hwndTabs, &strip);
        contentRect.bottom = std::max(contentRect.bottom, strip.bottom + ScaleTabsPx(1));
    }

    const COLORREF itemBg = selected ? palette.activeBg : (hovered ? palette.hoverBg : palette.inactiveBg);
    DrawRoundedRectDc(hdc, contentRect, selected ? ScaleTabsPx(4) : ScaleTabsPx(5), itemBg, selected ? palette.borderColor : itemBg);
    if (selected)
    {
        // Cover tab bottom stroke so active tab and editor page read as one continuous surface.
        RECT joinRect = contentRect;
        joinRect.top = std::max(joinRect.top, joinRect.bottom - std::max(1, ScaleTabsPx(2)));
        FillSolidRectDc(hdc, joinRect, itemBg);
    }

    HFONT drawFont = selected ? g_hTabFontActive : g_hTabFontRegular;
    HGDIOBJ oldFont = nullptr;
    if (drawFont)
        oldFont = SelectObject(hdc, drawFont);

    RECT textRect = {};
    if (showClose)
    {
        textRect = TabTextRect(contentRect);
    }
    else
    {
        textRect = contentRect;
        textRect.left += ScaleTabsPx(10);
        textRect.right -= ScaleTabsPx(10);
        if (textRect.right < textRect.left)
            textRect.right = textRect.left;
    }
    std::wstring label = DocumentTabLabel(g_documents[index]);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, selected ? palette.activeTextColor : palette.textColor);
    DrawTextW(hdc, label.c_str(), -1, &textRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
    if (oldFont)
        SelectObject(hdc, oldFont);

    if (showClose)
    {
        RECT closeRect = TabCloseRect(contentRect);
        if (closeHovered)
            DrawRoundedRectDc(hdc, closeRect, ScaleTabsPx(4), palette.closeHoverBg, palette.closeHoverBg);
        DrawCloseGlyph(hdc, closeRect, closeHovered ? palette.closeHoverFg : palette.closeColor);
    }
}

static void PaintTabStripVisual(HDC hdc)
{
    if (!hdc || !g_hwndTabs)
        return;

    const TabPaintPalette palette = GetTabPaintPalette(IsDarkMode());
    RECT rcClient{};
    GetClientRect(g_hwndTabs, &rcClient);
    DrawTabStripBackground(hdc, rcClient, palette);

    const int count = std::min(TabCtrl_GetItemCount(g_hwndTabs), static_cast<int>(g_documents.size()));
    for (int i = 0; i < count; ++i)
    {
        RECT itemRect{};
        if (TabCtrl_GetItemRect(g_hwndTabs, i, &itemRect))
            DrawTabItemVisual(hdc, i, itemRect, palette);
    }
}

static void UpdateTabsHoverState(HWND hwnd, POINT ptClient)
{
    TCHITTESTINFO hit{};
    hit.pt = ptClient;
    const int hoverIndex = TabCtrl_HitTest(hwnd, &hit);
    const bool hoverClose = hoverIndex >= 0 && IsTabCloseRectHit(hoverIndex, ptClient);

    if (hoverIndex != g_hoverTabIndex || hoverClose != g_hoverTabClose)
    {
        g_hoverTabIndex = hoverIndex;
        g_hoverTabClose = hoverClose;
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    if (!g_trackingTabsMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize = sizeof(tme);
        tme.dwFlags = TME_LEAVE;
        tme.hwndTrack = hwnd;
        if (TrackMouseEvent(&tme))
            g_trackingTabsMouse = true;
    }
}

static LRESULT HandleTabsCustomDraw(LPNMCUSTOMDRAW draw)
{
    if (!draw || !g_hwndTabs)
        return CDRF_DODEFAULT;

    g_tabsCustomDrawObserved = true;

    if (draw->dwDrawStage == CDDS_PREPAINT)
    {
        const TabPaintPalette palette = GetTabPaintPalette(IsDarkMode());
        RECT rcClient{};
        GetClientRect(g_hwndTabs, &rcClient);
        DrawTabStripBackground(draw->hdc, rcClient, palette);
        return CDRF_NOTIFYITEMDRAW;
    }

    if (draw->dwDrawStage != CDDS_ITEMPREPAINT)
        return CDRF_DODEFAULT;

    const int index = static_cast<int>(draw->dwItemSpec);
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return CDRF_SKIPDEFAULT;

    const TabPaintPalette palette = GetTabPaintPalette(IsDarkMode());
    DrawTabItemVisual(draw->hdc, index, draw->rc, palette);
    return CDRF_SKIPDEFAULT;
}

static LRESULT CALLBACK TabsSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_THEMECHANGED:
    case WM_STYLECHANGED:
    case WM_SETTINGCHANGE:
    case WM_DPICHANGED:
        RefreshTabsDpi();
        RefreshTabsVisualMetrics();
        g_tabsCustomDrawObserved = false;
        InvalidateRect(hwnd, nullptr, TRUE);
        break;
    case WM_PAINT:
    {
        const LRESULT result = CallWindowProcW(g_origTabsProc, hwnd, msg, wParam, lParam);
        if (!g_tabsCustomDrawObserved)
        {
            HDC hdc = GetDC(hwnd);
            if (hdc)
            {
                PaintTabStripVisual(hdc);
                ReleaseDC(hwnd, hdc);
            }
        }
        return result;
    }
    case WM_LBUTTONDOWN:
    {
        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        TCHITTESTINFO hit{};
        hit.pt = pt;
        int index = TabCtrl_HitTest(hwnd, &hit);
        if (index >= 0 && !IsTabCloseHotspot(index, pt))
        {
            g_draggingTab = true;
            g_dragTabIndex = index;
            SetCapture(hwnd);
        }
        break;
    }
    case WM_MOUSEMOVE:
    {
        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        UpdateTabsHoverState(hwnd, pt);

        if (!g_draggingTab || GetCapture() != hwnd)
            break;

        TCHITTESTINFO hit{};
        hit.pt = pt;
        int hoverIndex = TabCtrl_HitTest(hwnd, &hit);
        if (hoverIndex >= 0 && hoverIndex != g_dragTabIndex)
        {
            ReorderDocumentTab(g_dragTabIndex, hoverIndex);
            g_dragTabIndex = hoverIndex;
        }
        break;
    }
    case WM_MOUSELEAVE:
    {
        g_trackingTabsMouse = false;
        if (g_hoverTabIndex != -1 || g_hoverTabClose)
        {
            g_hoverTabIndex = -1;
            g_hoverTabClose = false;
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        break;
    }
    case WM_LBUTTONUP:
    {
        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        TCHITTESTINFO hit{};
        hit.pt = pt;
        int index = TabCtrl_HitTest(hwnd, &hit);

        if (g_draggingTab)
        {
            g_draggingTab = false;
            g_dragTabIndex = -1;
            if (GetCapture() == hwnd)
                ReleaseCapture();
        }

        if (index >= 0 && IsTabCloseHotspot(index, pt))
        {
            CloseDocumentTabAt(index);
            return 0;
        }
        break;
    }
    case WM_MBUTTONUP:
    {
        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        TCHITTESTINFO hit{};
        hit.pt = pt;
        int index = TabCtrl_HitTest(hwnd, &hit);
        if (index >= 0)
        {
            CloseDocumentTabAt(index);
            return 0;
        }
        break;
    }
    case WM_CAPTURECHANGED:
        g_draggingTab = false;
        g_dragTabIndex = -1;
        break;
    }
    return CallWindowProcW(g_origTabsProc, hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        g_hwndMain = hwnd;
        RefreshTabsDpi();
        DragAcceptFiles(hwnd, TRUE);
        const wchar_t *richEditClass = nullptr;
        if (g_hRichEditModule)
        {
            FreeLibrary(g_hRichEditModule);
            g_hRichEditModule = nullptr;
        }
        g_hRichEditModule = LoadLibraryW(L"Msftedit.dll");
        if (g_hRichEditModule)
        {
            richEditClass = MSFTEDIT_CLASS;
        }
        else
        {
            g_hRichEditModule = LoadLibraryW(L"Riched20.dll");
            if (g_hRichEditModule)
            {
                richEditClass = RICHEDIT_CLASSW;
            }
            else
            {
                MessageBoxW(hwnd,
                            L"Cannot load RichEdit control.\n",
                            L"Error", MB_ICONERROR | MB_OK);
                return -1;
            }
        }
        g_editorClassName = richEditClass;
        DWORD editorStyle = WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN | ES_NOHIDESEL;
        if (!(g_state.wordWrap && !g_state.largeFileMode))
            editorStyle |= WS_HSCROLL | ES_AUTOHSCROLL;
        g_hwndEditor = CreateWindowExW(0, richEditClass, nullptr,
                                       editorStyle,
                                       0, 0, 100, 100, hwnd, reinterpret_cast<HMENU>(IDC_EDITOR), GetModuleHandleW(nullptr), nullptr);
        g_origEditorProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndEditor, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(EditorSubclassProc)));
        ConfigureEditorControl(g_hwndEditor);
        g_hwndTabs = CreateWindowExW(0, WC_TABCONTROLW, L"",
                                     WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_TABSTOP | TCS_HOTTRACK | TCS_FOCUSNEVER,
                                     0, 0, 100, ScaleTabsPx(32), hwnd, reinterpret_cast<HMENU>(IDC_TABS), GetModuleHandleW(nullptr), nullptr);
        if (g_hwndTabs)
        {
            RefreshTabsDpi();
            RefreshTabsVisualMetrics();
            SetWindowTheme(g_hwndTabs, L"Explorer", nullptr);
            g_tabsCustomDrawObserved = false;
            g_origTabsProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndTabs, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(TabsSubclassProc)));
        }
        g_hwndStatus = CreateWindowExW(0, STATUSCLASSNAMEW, nullptr,
                                       WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDC_STATUSBAR), GetModuleHandleW(nullptr), nullptr);
        g_origStatusProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndStatus, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(StatusSubclassProc)));
        SendMessageW(g_hwndEditor, EM_EXLIMITTEXT, 0, static_cast<LPARAM>(-1));
        SendMessageW(g_hwndEditor, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
        ApplyFont();
        SetupStatusBarParts();
        UpdateMenuStrings();
        UpdateRecentFilesMenu();
        UpdateLanguageMenu();
        HMENU hMainMenu = GetMenu(g_hwndMain);
        if (hMainMenu)
        {
            CheckMenuItem(hMainMenu, IDM_FORMAT_WORDWRAP, g_state.wordWrap ? MF_CHECKED : MF_UNCHECKED);
            CheckMenuItem(hMainMenu, IDM_VIEW_STATUSBAR, g_state.showStatusBar ? MF_CHECKED : MF_UNCHECKED);
            CheckMenuItem(hMainMenu, IDM_VIEW_ALWAYSONTOP, g_state.alwaysOnTop ? MF_CHECKED : MF_UNCHECKED);
        }
        if (g_state.alwaysOnTop)
            SetWindowPos(g_hwndMain, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        if (g_state.windowOpacity < 255)
        {
            SetWindowLongW(g_hwndMain, GWL_EXSTYLE, GetWindowLongW(g_hwndMain, GWL_EXSTYLE) | WS_EX_LAYERED);
            SetLayeredWindowAttributes(g_hwndMain, 0, g_state.windowOpacity, LWA_ALPHA);
        }
        if (!g_state.customIconPath.empty())
        {
            if (!ApplyCustomIcon(g_state.customIconPath, g_state.customIconIndex, false))
            {
                g_state.customIconPath.clear();
                g_state.customIconIndex = 0;
            }
        }
        if (g_state.background.enabled && !g_state.background.imagePath.empty())
        {
            LoadBackgroundImage(g_state.background.imagePath);
            SetBackgroundPosition(g_state.background.position);
        }
        CreateInitialDocumentTabIfNeeded();
        UpdateRuntimeMenuStates();
        UpdateTitle();
        ResizeControls();
        UpdateStatus();
        ApplyTheme();
        SetFocus(g_hwndEditor);
        return 0;
    }
    case WM_UAHDRAWMENU:
    {
        if (IsDarkMode())
        {
            UAHMENU *pUDM = reinterpret_cast<UAHMENU *>(lParam);
            MENUBARINFO mbi = {};
            mbi.cbSize = sizeof(mbi);
            if (GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbi))
            {
                RECT rcWindow;
                GetWindowRect(hwnd, &rcWindow);
                RECT rcMenuBar = mbi.rcBar;
                OffsetRect(&rcMenuBar, -rcWindow.left, -rcWindow.top);
                const COLORREF menuBg = ThemeColorMenuBackground(true);
                const COLORREF borderColor = ThemeColorChromeBorder(true);
                HBRUSH hbrMenu = g_hbrMenuDark ? g_hbrMenuDark : CreateSolidBrush(menuBg);
                FillRect(pUDM->hdc, &rcMenuBar, hbrMenu ? hbrMenu : reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
                if (!g_hbrMenuDark && hbrMenu)
                    DeleteObject(hbrMenu);

                RECT separator = rcMenuBar;
                separator.top = std::max(separator.top, separator.bottom - 1);
                HBRUSH hbrBorder = CreateSolidBrush(borderColor);
                if (hbrBorder)
                {
                    FillRect(pUDM->hdc, &separator, hbrBorder);
                    DeleteObject(hbrBorder);
                }
            }
            return TRUE;
        }
        break;
    }
    case WM_UAHDRAWMENUITEM:
    {
        if (IsDarkMode())
        {
            UAHDRAWMENUITEM *pUDMI = reinterpret_cast<UAHDRAWMENUITEM *>(lParam);
            wchar_t szText[256] = {};
            MENUITEMINFOW mii = {};
            mii.cbSize = sizeof(mii);
            mii.fMask = MIIM_STRING;
            mii.dwTypeData = szText;
            mii.cch = 255;
            GetMenuItemInfoW(pUDMI->um.hMenu, pUDMI->umi.iPosition, TRUE, &mii);
            COLORREF bgColor = ThemeColorMenuBackground(true);
            COLORREF textColor = ThemeColorMenuText(true);
            if ((pUDMI->dis.itemState & ODS_HOTLIGHT) || (pUDMI->dis.itemState & ODS_SELECTED))
                bgColor = ThemeColorMenuHoverBackground(true);
            if (pUDMI->dis.itemState & ODS_DISABLED)
                textColor = RGB(144, 150, 160);
            HBRUSH hbr = CreateSolidBrush(bgColor);
            FillRect(pUDMI->um.hdc, &pUDMI->dis.rcItem, hbr);
            DeleteObject(hbr);
            NONCLIENTMETRICSW ncm = {};
            ncm.cbSize = sizeof(ncm);
            SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
            HFONT hFont = CreateFontIndirectW(&ncm.lfMenuFont);
            HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(pUDMI->um.hdc, hFont));
            SetBkMode(pUDMI->um.hdc, TRANSPARENT);
            SetTextColor(pUDMI->um.hdc, textColor);
            RECT rcText = pUDMI->dis.rcItem;
            DrawTextW(pUDMI->um.hdc, szText, -1, &rcText, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
            SelectObject(pUDMI->um.hdc, hOldFont);
            DeleteObject(hFont);
            return TRUE;
        }
        break;
    }
    case WM_NCPAINT:
    case WM_NCACTIVATE:
    {
        LRESULT result = DefWindowProcW(hwnd, msg, wParam, lParam);
        if (IsDarkMode())
        {
            HDC hdc = GetWindowDC(hwnd);
            MENUBARINFO mbi = {};
            mbi.cbSize = sizeof(mbi);
            if (GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbi))
            {
                RECT rcWindow;
                GetWindowRect(hwnd, &rcWindow);
                RECT rcMenuBar = mbi.rcBar;
                OffsetRect(&rcMenuBar, -rcWindow.left, -rcWindow.top);
                rcMenuBar.bottom += 1;
                const COLORREF menuBg = ThemeColorMenuBackground(true);
                const COLORREF menuText = ThemeColorMenuText(true);
                const COLORREF borderColor = ThemeColorChromeBorder(true);
                HBRUSH hbrMenu = g_hbrMenuDark ? g_hbrMenuDark : CreateSolidBrush(menuBg);
                FillRect(hdc, &rcMenuBar, hbrMenu ? hbrMenu : reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
                if (!g_hbrMenuDark && hbrMenu)
                    DeleteObject(hbrMenu);

                RECT separator = rcMenuBar;
                separator.top = std::max(separator.top, separator.bottom - 1);
                HBRUSH hbrBorder = CreateSolidBrush(borderColor);
                if (hbrBorder)
                {
                    FillRect(hdc, &separator, hbrBorder);
                    DeleteObject(hbrBorder);
                }
                HMENU hMenu = GetMenu(hwnd);
                int itemCount = GetMenuItemCount(hMenu);
                NONCLIENTMETRICSW ncm = {};
                ncm.cbSize = sizeof(ncm);
                SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
                HFONT hFont = CreateFontIndirectW(&ncm.lfMenuFont);
                HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(hdc, hFont));
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, menuText);
                for (int i = 0; i < itemCount; i++)
                {
                    RECT rcItem;
                    if (GetMenuBarInfo(hwnd, OBJID_MENU, i + 1, &mbi))
                    {
                        rcItem = mbi.rcBar;
                        OffsetRect(&rcItem, -rcWindow.left, -rcWindow.top);
                        wchar_t szText[256] = {};
                        MENUITEMINFOW mii = {};
                        mii.cbSize = sizeof(mii);
                        mii.fMask = MIIM_STRING;
                        mii.dwTypeData = szText;
                        mii.cch = 255;
                        GetMenuItemInfoW(hMenu, i, TRUE, &mii);
                        DrawTextW(hdc, szText, -1, &rcItem, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
                    }
                }
                SelectObject(hdc, hOldFont);
                DeleteObject(hFont);
            }
            ReleaseDC(hwnd, hdc);
        }
        return result;
    }
    case WM_SETTINGCHANGE:
    {
        if (lParam && wcscmp(reinterpret_cast<LPCWSTR>(lParam), L"ImmersiveColorSet") == 0)
            ApplyTheme();
        return 0;
    }
    case WM_DROPFILES:
    {
        HDROP hDrop = reinterpret_cast<HDROP>(wParam);
        UINT fileCount = DragQueryFileW(hDrop, 0xFFFFFFFF, nullptr, 0);
        if (!g_state.useTabs)
        {
            wchar_t path[MAX_PATH] = {};
            if (fileCount > 0 && DragQueryFileW(hDrop, fileCount - 1, path, MAX_PATH))
                OpenPathInTabs(path);
        }
        else
        {
            for (UINT i = 0; i < fileCount; ++i)
            {
                wchar_t path[MAX_PATH] = {};
                if (DragQueryFileW(hDrop, i, path, MAX_PATH))
                    OpenPathInTabs(path);
            }
        }
        DragFinish(hDrop);
        return 0;
    }
    case WM_SIZE:
        RefreshTabsDpi();
        ResizeControls();
        UpdateStatus();
        return 0;
    case WM_TIMER:
        if (wParam == kSessionAutosaveTimerId)
        {
            if (g_sessionDirty && !g_state.closing)
            {
                if (g_sessionRetryAtTick != 0)
                {
                    const DWORD now = GetTickCount();
                    if (static_cast<int32_t>(now - g_sessionRetryAtTick) < 0)
                        return 0;
                }

                bool persisted = true;
                if (g_state.startupBehavior == StartupBehavior::ResumeAll)
                    persisted = SaveOpenDocumentSession(false);
                if (persisted)
                {
                    g_sessionDirty = false;
                    g_sessionRetryAtTick = 0;
                }
                else
                {
                    g_sessionRetryAtTick = GetTickCount() + kSessionRetryBackoffMs;
                }
            }
            return 0;
        }
        break;
    case WM_SETFOCUS:
        SetFocus(g_hwndEditor);
        return 0;
    case WM_CTLCOLOREDIT:
        if (g_state.background.enabled && g_bgImage && reinterpret_cast<HWND>(lParam) == g_hwndEditor)
        {
            SetBkMode(reinterpret_cast<HDC>(wParam), TRANSPARENT);
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }
        break;
    case WM_CTLCOLORSTATIC:
        if (reinterpret_cast<HWND>(lParam) == g_hwndStatus && IsDarkMode())
        {
            HDC hdc = reinterpret_cast<HDC>(wParam);
            SetTextColor(hdc, ThemeColorStatusText(true));
            SetBkColor(hdc, ThemeColorStatusBackground(true));
            return reinterpret_cast<LRESULT>(g_hbrStatusDark ? g_hbrStatusDark : GetStockObject(BLACK_BRUSH));
        }
        break;
    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT pDIS = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);
        if (pDIS->hwndItem == g_hwndStatus && IsDarkMode())
        {
            const COLORREF bgColor = ThemeColorStatusBackground(true);
            const COLORREF textColor = ThemeColorStatusText(true);
            const COLORREF borderColor = ThemeColorChromeBorder(true);
            HBRUSH hbr = g_hbrStatusDark ? g_hbrStatusDark : CreateSolidBrush(bgColor);
            FillRect(pDIS->hDC, &pDIS->rcItem, hbr);
            if (!g_hbrStatusDark)
                DeleteObject(hbr);
            SetBkMode(pDIS->hDC, TRANSPARENT);
            SetTextColor(pDIS->hDC, textColor);
            int part = static_cast<int>(pDIS->itemID);
            if (part >= 0 && part < 4)
            {
                RECT rc = pDIS->rcItem;
                rc.left += 6;
                DrawTextW(pDIS->hDC, g_statusTexts[part].c_str(), -1, &rc, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);
                if (part < 3)
                {
                    RECT separator = pDIS->rcItem;
                    separator.left = separator.right - 1;
                    HBRUSH hbrSep = CreateSolidBrush(borderColor);
                    if (hbrSep)
                    {
                        FillRect(pDIS->hDC, &separator, hbrSep);
                        DeleteObject(hbrSep);
                    }
                }
            }
            return TRUE;
        }
        break;
    }
    case WM_CONTEXTMENU:
    {
        if (reinterpret_cast<HWND>(wParam) != g_hwndEditor)
            break;

        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        if (pt.x == -1 && pt.y == -1)
        {
            DWORD start = 0;
            SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), 0);
            POINTL charPos = {};
            if (SendMessageW(g_hwndEditor, EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&charPos), start) != -1)
            {
                pt.x = static_cast<LONG>(charPos.x);
                pt.y = static_cast<LONG>(charPos.y);
                ClientToScreen(g_hwndEditor, &pt);
            }
            else
            {
                GetCursorPos(&pt);
            }
        }

        HMENU hPopup = CreatePopupMenu();
        if (!hPopup)
            return 0;

        const auto &lang = GetLangStrings();
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_UNDO, MenuLabelForContext(lang.menuUndo).c_str());
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_REDO, MenuLabelForContext(lang.menuRedo).c_str());
        AppendMenuW(hPopup, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_CUT, MenuLabelForContext(lang.menuCut).c_str());
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_COPY, MenuLabelForContext(lang.menuCopy).c_str());
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_PASTE, MenuLabelForContext(lang.menuPaste).c_str());
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_DELETE, MenuLabelForContext(lang.menuDelete).c_str());
        AppendMenuW(hPopup, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(hPopup, MF_STRING, IDM_EDIT_SELECTALL, MenuLabelForContext(lang.menuSelectAll).c_str());

        DWORD selStart = 0, selEnd = 0;
        SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&selStart), reinterpret_cast<LPARAM>(&selEnd));
        const bool hasSelection = selStart != selEnd;
        const bool canUndo = SendMessageW(g_hwndEditor, EM_CANUNDO, 0, 0) != 0;
        const bool canRedo = SendMessageW(g_hwndEditor, EM_CANREDO, 0, 0) != 0;
        const bool canPaste = IsClipboardFormatAvailable(CF_UNICODETEXT) || IsClipboardFormatAvailable(CF_TEXT);

        EnableMenuItem(hPopup, IDM_EDIT_UNDO, MF_BYCOMMAND | (canUndo ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(hPopup, IDM_EDIT_REDO, MF_BYCOMMAND | (canRedo ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(hPopup, IDM_EDIT_CUT, MF_BYCOMMAND | (hasSelection ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(hPopup, IDM_EDIT_COPY, MF_BYCOMMAND | (hasSelection ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(hPopup, IDM_EDIT_PASTE, MF_BYCOMMAND | (canPaste ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(hPopup, IDM_EDIT_DELETE, MF_BYCOMMAND | (hasSelection ? MF_ENABLED : MF_GRAYED));

        const UINT cmd = TrackPopupMenuLightweight(hPopup, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, hwnd);
        DestroyMenu(hPopup);
        if (cmd != 0)
            SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
        return 0;
    }
    case WM_COMMAND:
    {
        WORD cmd = LOWORD(wParam);
        if (cmd == IDC_EDITOR && HIWORD(wParam) == EN_CHANGE)
        {
            if (g_switchingDocument)
                return 0;
            if (!g_state.modified)
            {
                g_state.modified = true;
                UpdateTitle();
            }
            UpdateStatus();
            SyncDocumentFromState(g_activeDocument, false);
            MarkSessionDirty();
            return 0;
        }

        if (cmd >= IDM_FILE_RECENT_BASE && cmd < IDM_FILE_RECENT_BASE + MAX_RECENT_FILES)
        {
            int idx = cmd - IDM_FILE_RECENT_BASE;
            if (idx < static_cast<int>(g_state.recentFiles.size()))
                OpenPathInTabs(g_state.recentFiles[idx]);
            return 0;
        }
        switch (cmd)
        {
        case IDM_FILE_NEW:
            if (g_state.useTabs)
            {
                CreateNewDocumentTab();
            }
            else
            {
                FileNew();
                EnsureSingleDocumentModel(true);
                UpdateRuntimeMenuStates();
            }
            break;
        case IDM_FILE_OPEN:
            if (g_state.useTabs)
            {
                OpenFileInNewDocumentTabDialog();
            }
            else
            {
                FileOpen();
                EnsureSingleDocumentModel(true);
                UpdateRuntimeMenuStates();
            }
            break;
        case IDM_FILE_CLOSETAB:
            if (g_state.useTabs)
                CloseCurrentDocumentTab();
            break;
        case IDM_FILE_REOPENCLOSEDTAB:
            if (g_state.useTabs)
                ReopenClosedDocumentTab();
            break;
        case IDM_FILE_NEXTTAB:
            if (g_state.useTabs)
                SwitchToNextDocumentTab(false);
            break;
        case IDM_FILE_PREVTAB:
            if (g_state.useTabs)
                SwitchToNextDocumentTab(true);
            break;
        case IDM_FILE_SAVE:
            FileSave();
            break;
        case IDM_FILE_SAVEAS:
            FileSaveAs();
            break;
        case IDM_FILE_PRINT:
            FilePrint();
            break;
        case IDM_FILE_PAGESETUP:
            FilePageSetup();
            break;
        case IDM_FILE_EXIT:
            SendMessageW(hwnd, WM_CLOSE, 0, 0);
            break;
        case IDM_EDIT_UNDO:
            EditUndo();
            break;
        case IDM_EDIT_REDO:
            EditRedo();
            break;
        case IDM_EDIT_CUT:
            EditCut();
            break;
        case IDM_EDIT_COPY:
            EditCopy();
            break;
        case IDM_EDIT_PASTE:
            EditPaste();
            break;
        case IDM_EDIT_DELETE:
            EditDelete();
            break;
        case IDM_EDIT_FIND:
            EditFind();
            break;
        case IDM_EDIT_FINDNEXT:
            EditFindNext();
            break;
        case IDM_EDIT_FINDPREV:
            EditFindPrev();
            break;
        case IDM_EDIT_REPLACE:
            EditReplace();
            break;
        case IDM_EDIT_GOTO:
            EditGoto();
            break;
        case IDM_EDIT_SELECTALL:
            EditSelectAll();
            break;
        case IDM_EDIT_TIMEDATE:
            EditTimeDate();
            break;
        case IDM_FORMAT_WORDWRAP:
            if (!g_state.largeFileMode)
                FormatWordWrap();
            break;
        case IDM_FORMAT_FONT:
            FormatFont();
            break;
        case IDM_VIEW_ZOOMIN:
            ViewZoomIn();
            break;
        case IDM_VIEW_ZOOMOUT:
            ViewZoomOut();
            break;
        case IDM_VIEW_ZOOMDEFAULT:
            ViewZoomDefault();
            break;
        case IDM_VIEW_STATUSBAR:
            ViewStatusBar();
            break;
        case IDM_VIEW_DARKMODE:
            ToggleDarkMode();
            break;
        case IDM_VIEW_TRANSPARENCY:
            ViewTransparency();
            break;
        case IDM_VIEW_ALWAYSONTOP:
            ViewAlwaysOnTop();
            break;
        case IDM_VIEW_USETABS:
            if (g_state.useTabs && g_documents.size() > 1)
            {
                const auto &lang = GetLangStrings();
                MessageBoxW(hwnd,
                            L"Close other tabs first, then disable tab mode.",
                            lang.appName.c_str(),
                            MB_OK | MB_ICONINFORMATION);
                break;
            }
            ApplyTabsMode(!g_state.useTabs);
            SaveFontSettings();
            break;
        case IDM_VIEW_STARTUP_CLASSIC:
            g_state.startupBehavior = StartupBehavior::Classic;
            UpdateRuntimeMenuStates();
            UpdateSessionAutosaveTimer();
            SaveOpenDocumentSession(true);
            SaveFontSettings();
            g_sessionDirty = false;
            break;
        case IDM_VIEW_STARTUP_RESUMEALL:
            g_state.startupBehavior = StartupBehavior::ResumeAll;
            UpdateRuntimeMenuStates();
            UpdateSessionAutosaveTimer();
            SaveOpenDocumentSession(true);
            SaveFontSettings();
            g_sessionDirty = false;
            break;
        case IDM_VIEW_STARTUP_RESUMESAVED:
            g_state.startupBehavior = StartupBehavior::ResumeSaved;
            UpdateRuntimeMenuStates();
            UpdateSessionAutosaveTimer();
            SaveOpenDocumentSession(true);
            SaveFontSettings();
            g_sessionDirty = false;
            break;
        case IDM_VIEW_BG_SELECT:
            ViewSelectBackground();
            break;
        case IDM_VIEW_BG_CLEAR:
            ViewClearBackground();
            break;
        case IDM_VIEW_BG_OPACITY:
            ViewBackgroundOpacity();
            break;
        case IDM_VIEW_BG_POS_TOPLEFT:
            SetBackgroundPosition(BgPosition::TopLeft);
            break;
        case IDM_VIEW_BG_POS_TOPCENTER:
            SetBackgroundPosition(BgPosition::TopCenter);
            break;
        case IDM_VIEW_BG_POS_TOPRIGHT:
            SetBackgroundPosition(BgPosition::TopRight);
            break;
        case IDM_VIEW_BG_POS_CENTERLEFT:
            SetBackgroundPosition(BgPosition::CenterLeft);
            break;
        case IDM_VIEW_BG_POS_CENTER:
            SetBackgroundPosition(BgPosition::Center);
            break;
        case IDM_VIEW_BG_POS_CENTERRIGHT:
            SetBackgroundPosition(BgPosition::CenterRight);
            break;
        case IDM_VIEW_BG_POS_BOTTOMLEFT:
            SetBackgroundPosition(BgPosition::BottomLeft);
            break;
        case IDM_VIEW_BG_POS_BOTTOMCENTER:
            SetBackgroundPosition(BgPosition::BottomCenter);
            break;
        case IDM_VIEW_BG_POS_BOTTOMRIGHT:
            SetBackgroundPosition(BgPosition::BottomRight);
            break;
        case IDM_VIEW_BG_POS_TILE:
            SetBackgroundPosition(BgPosition::Tile);
            break;
        case IDM_VIEW_BG_POS_STRETCH:
            SetBackgroundPosition(BgPosition::Stretch);
            break;
        case IDM_VIEW_BG_POS_FIT:
            SetBackgroundPosition(BgPosition::Fit);
            break;
        case IDM_VIEW_BG_POS_FILL:
            SetBackgroundPosition(BgPosition::Fill);
            break;
        case IDM_VIEW_LANG_EN:
            if (g_hwndFindDlg)
            {
                DestroyWindow(g_hwndFindDlg);
                g_hwndFindDlg = nullptr;
            }
            SetLanguage(LangID::EN);
            UpdateMenuStrings();
            UpdateRecentFilesMenu();
            UpdateLanguageMenu();
            RefreshAllDocumentTabLabels();
            UpdateTitle();
            UpdateStatus();
            UpdateRuntimeMenuStates();
            break;
        case IDM_VIEW_LANG_JA:
            if (g_hwndFindDlg)
            {
                DestroyWindow(g_hwndFindDlg);
                g_hwndFindDlg = nullptr;
            }
            SetLanguage(LangID::JA);
            UpdateMenuStrings();
            UpdateRecentFilesMenu();
            UpdateLanguageMenu();
            RefreshAllDocumentTabLabels();
            UpdateTitle();
            UpdateStatus();
            UpdateRuntimeMenuStates();
            break;
        case IDM_VIEW_ICON_CHANGE:
            ViewChangeIcon();
            break;
        case IDM_VIEW_ICON_SYSTEM:
            ViewChooseSystemIcon();
            break;
        case IDM_VIEW_ICON_RESET:
            ViewResetIcon();
            break;
        case IDM_HELP_CHECKUPDATES:
            HelpCheckUpdates();
            break;
        case IDM_HELP_PERF_BENCHMARK:
            HelpRunPerformanceBenchmark();
            break;
        case IDM_HELP_ABOUT:
            HelpAbout();
            break;
        }
        SyncDocumentFromState(g_activeDocument, false);
        MarkSessionDirty();
        return 0;
    }
    case WM_NOTIFY:
    {
        NMHDR *pnmh = reinterpret_cast<NMHDR *>(lParam);
        if (pnmh->hwndFrom == g_hwndTabs && !g_state.useTabs)
            return 0;
        if (pnmh->hwndFrom == g_hwndTabs && pnmh->code == NM_CUSTOMDRAW)
            return HandleTabsCustomDraw(reinterpret_cast<LPNMCUSTOMDRAW>(lParam));
        if (pnmh->hwndFrom == g_hwndTabs && pnmh->code == TCN_SELCHANGE)
        {
            if (g_updatingTabs)
                return 0;
            int index = TabCtrl_GetCurSel(g_hwndTabs);
            SwitchToDocument(index);
            return 0;
        }
        if (pnmh->hwndFrom == g_hwndTabs && pnmh->code == NM_RCLICK)
        {
            POINT pt{};
            GetCursorPos(&pt);
            POINT local = pt;
            ScreenToClient(g_hwndTabs, &local);
            TCHITTESTINFO hit{};
            hit.pt = local;
            int index = TabCtrl_HitTest(g_hwndTabs, &hit);
            if (index >= 0)
                SwitchToDocument(index);

            HMENU hPopup = CreatePopupMenu();
            if (!hPopup)
                return 0;

            const auto &lang = GetLangStrings();
            AppendMenuW(hPopup, MF_STRING, IDM_FILE_NEW, MenuLabelForContext(lang.menuNew).c_str());
            AppendMenuW(hPopup, MF_STRING, IDM_FILE_CLOSETAB, L"Close Tab");
            AppendMenuW(hPopup, MF_STRING, IDM_FILE_REOPENCLOSEDTAB, L"Reopen Closed Tab");
            if (g_documents.size() <= 1)
                EnableMenuItem(hPopup, IDM_FILE_CLOSETAB, MF_BYCOMMAND | MF_GRAYED);
            if (g_closedDocuments.empty())
                EnableMenuItem(hPopup, IDM_FILE_REOPENCLOSEDTAB, MF_BYCOMMAND | MF_GRAYED);

            UINT cmd = TrackPopupMenuLightweight(hPopup, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD,
                                                 pt.x, pt.y, hwnd);
            DestroyMenu(hPopup);
            if (cmd != 0)
                SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
            return 0;
        }
        if (pnmh->hwndFrom == g_hwndStatus && pnmh->code == NM_CUSTOMDRAW)
        {
            if (IsDarkMode())
            {
                LPNMCUSTOMDRAW lpnmcd = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
                const COLORREF bgColor = ThemeColorStatusBackground(true);
                const COLORREF textColor = ThemeColorStatusText(true);
                const COLORREF borderColor = ThemeColorChromeBorder(true);
                if (lpnmcd->dwDrawStage == CDDS_PREPAINT)
                    return CDRF_NOTIFYITEMDRAW;
                if (lpnmcd->dwDrawStage == CDDS_ITEMPREPAINT)
                {
                    HBRUSH hbr = g_hbrStatusDark ? g_hbrStatusDark : CreateSolidBrush(bgColor);
                    FillRect(lpnmcd->hdc, &lpnmcd->rc, hbr);
                    if (!g_hbrStatusDark && hbr)
                        DeleteObject(hbr);
                    SetBkMode(lpnmcd->hdc, TRANSPARENT);
                    SetBkColor(lpnmcd->hdc, bgColor);
                    SetTextColor(lpnmcd->hdc, textColor);
                    wchar_t buf[256] = {};
                    int part = static_cast<int>(lpnmcd->dwItemSpec);
                    SendMessageW(g_hwndStatus, SB_GETTEXTW, part, reinterpret_cast<LPARAM>(buf));
                    RECT rc = lpnmcd->rc;
                    rc.left += 6;
                    DrawTextW(lpnmcd->hdc, buf, -1, &rc, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);
                    if (part < 3)
                    {
                        RECT separator = lpnmcd->rc;
                        separator.left = separator.right - 1;
                        HBRUSH hbrSep = CreateSolidBrush(borderColor);
                        if (hbrSep)
                        {
                            FillRect(lpnmcd->hdc, &separator, hbrSep);
                            DeleteObject(hbrSep);
                        }
                    }
                    return CDRF_SKIPDEFAULT;
                }
            }
        }
        if (pnmh->hwndFrom == g_hwndEditor && pnmh->code == EN_SELCHANGE)
        {
            UpdateStatus();
        }
        if (pnmh->hwndFrom == g_hwndEditor && pnmh->code == EN_LINK)
            return 1;
        return 0;
    }
    case WM_CLOSE:
        if (g_state.closing)
            return 0;
        if (!ConfirmCloseForCurrentStartupBehavior())
            return 0;
        g_state.closing = true;
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
    {
        WINDOWPLACEMENT placement = {};
        placement.length = sizeof(placement);
        if (GetWindowPlacement(g_hwndMain, &placement))
        {
            RECT rc = placement.rcNormalPosition;
            const int width = rc.right - rc.left;
            const int height = rc.bottom - rc.top;
            if (width > 0 && height > 0)
            {
                g_state.windowX = rc.left;
                g_state.windowY = rc.top;
                g_state.windowWidth = width;
                g_state.windowHeight = height;
            }
            g_state.windowMaximized = (placement.showCmd == SW_SHOWMAXIMIZED);
        }
        KillTimer(hwnd, kSessionAutosaveTimerId);
        SaveOpenDocumentSession(true);
        g_sessionDirty = false;
        SaveFontSettings();
        if (g_state.hFont)
        {
            DeleteObject(g_state.hFont);
            g_state.hFont = nullptr;
        }
        if (g_bgImage)
        {
            delete g_bgImage;
            g_bgImage = nullptr;
        }
        if (g_bgBitmap)
        {
            DeleteObject(g_bgBitmap);
            g_bgBitmap = nullptr;
        }
        if (g_hCustomIcon)
        {
            DestroyIcon(g_hCustomIcon);
            g_hCustomIcon = nullptr;
        }
        if (g_hbrStatusDark)
        {
            DeleteObject(g_hbrStatusDark);
            g_hbrStatusDark = nullptr;
        }
        if (g_hbrMenuDark)
        {
            DeleteObject(g_hbrMenuDark);
            g_hbrMenuDark = nullptr;
        }
        if (g_hbrDialogDark)
        {
            DeleteObject(g_hbrDialogDark);
            g_hbrDialogDark = nullptr;
        }
        if (g_hbrDialogEditDark)
        {
            DeleteObject(g_hbrDialogEditDark);
            g_hbrDialogEditDark = nullptr;
        }
        if (g_hRichEditModule)
        {
            FreeLibrary(g_hRichEditModule);
            g_hRichEditModule = nullptr;
        }
        DestroyTabFonts();
        PostQuitMessage(0);
        return 0;
    }
    case WM_MOUSEWHEEL:
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            int delta = GET_WHEEL_DELTA_WPARAM(wParam);
            if (delta > 0)
                ViewZoomIn();
            else
                ViewZoomOut();
            return 0;
        }
        break;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow)
{
    InitLanguage();
    typedef BOOL(WINAPI * fnSetProcessDpiAwarenessContext)(DPI_AWARENESS_CONTEXT);
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32)
    {
#ifdef __GNUC__
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wcast-function-type"
#endif
        auto setProcDPI = reinterpret_cast<fnSetProcessDpiAwarenessContext>(GetProcAddress(hUser32, "SetProcessDpiAwarenessContext"));
#ifdef __GNUC__
#pragma GCC diagnostic pop
#endif
        if (setProcDPI)
            setProcDPI(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    }

    LoadFontSettings();

    bool benchmarkOnly = false;
    std::vector<std::wstring> startupArgs;
    int argc = 0;
    LPWSTR *argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv)
    {
        for (int i = 1; i < argc; ++i)
        {
            if (!argv[i] || argv[i][0] == L'\0')
                continue;

            if (lstrcmpiW(argv[i], L"--benchmark-ci") == 0 || lstrcmpiW(argv[i], L"/benchmark-ci") == 0)
            {
                benchmarkOnly = true;
                continue;
            }

            startupArgs.emplace_back(argv[i]);
        }
        LocalFree(argv);
    }

    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_NOTEPAD));
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszMenuName = MAKEINTRESOURCEW(IDR_MAINMENU);
    wc.lpszClassName = L"NotepadClass";
    wc.hIconSm = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_NOTEPAD));
    RegisterClassExW(&wc);

    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    Gdiplus::GdiplusStartup(&g_gdiplusToken, &gdiplusStartupInput, nullptr);
    INITCOMMONCONTROLSEX icc = {sizeof(icc), ICC_BAR_CLASSES};
    InitCommonControlsEx(&icc);

    const auto &lang = GetLangStrings();
    std::wstring initialTitle = lang.untitled + L" - " + lang.appName;
    g_hwndMain = CreateWindowExW(0, L"NotepadClass", initialTitle.c_str(),
                                 WS_OVERLAPPEDWINDOW | WS_MAXIMIZEBOX, g_state.windowX, g_state.windowY, g_state.windowWidth, g_state.windowHeight,
                                 nullptr, nullptr, hInstance, nullptr);
    g_hAccel = LoadAcceleratorsW(hInstance, MAKEINTRESOURCEW(IDR_ACCEL));
    int showCmd = benchmarkOnly ? SW_HIDE : nCmdShow;
    if (!benchmarkOnly && g_state.windowMaximized && (nCmdShow == SW_SHOW || nCmdShow == SW_SHOWNORMAL || nCmdShow == SW_SHOWDEFAULT))
        showCmd = SW_SHOWMAXIMIZED;
    ShowWindow(g_hwndMain, showCmd);
    UpdateWindow(g_hwndMain);

    if (benchmarkOnly)
    {
        bool allPassed = false;
        bool allExecuted = false;
        const bool ok = RunPerformanceBenchmark(false, nullptr, &allPassed, &allExecuted);
        Gdiplus::GdiplusShutdown(g_gdiplusToken);
        return ok ? 0 : 1;
    }

    RestoreOpenDocumentSession();

    if (!startupArgs.empty())
    {
        if (!g_state.useTabs)
        {
            std::wstring startupPath;
            for (const auto &arg : startupArgs)
                startupPath = arg;
            if (!startupPath.empty())
                OpenPathInTabs(startupPath, true);
        }
        else
        {
            for (const auto &arg : startupArgs)
                OpenPathInTabs(arg, false);
        }
    }

    UpdateRuntimeMenuStates();
    UpdateSessionAutosaveTimer();

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0))
    {
        if (g_hwndFindDlg && IsDialogMessageW(g_hwndFindDlg, &msg))
            continue;
        if (!TranslateAcceleratorW(g_hwndMain, g_hAccel, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    Gdiplus::GdiplusShutdown(g_gdiplusToken);
    return static_cast<int>(msg.wParam);
}
