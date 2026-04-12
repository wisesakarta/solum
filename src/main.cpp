/*
  Technical Standard

  Main entry point and window procedure for Technical Standard Note text editor application.
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
#include "core/crash_diagnostics.h"
#include "core/file_dialog_filters.h"
#include "core/types.h"
#include "core/globals.h"
#include "modules/theme.h"
#include "modules/editor.h"
#include "modules/file.h"
#include "modules/ui.h"
#include "modules/background.h"
#include "modules/command_routing.h"
#include "modules/dialog.h"
#include "modules/commands.h"
#include "modules/settings.h"
#include "modules/menu.h"
#include "modules/tab_document.h"
#include "modules/tab_layout.h"
#include "modules/tab_model_ops.h"
#include "modules/tab_session_io.h"
#include "modules/design_system.h"
#include "modules/premium_orchestrator.h"
#include "modules/tab_spin_chrome.h"
#include "modules/custom_scrollbar.h"
#include "modules/selection_aura.h"
#include "modules/command_palette.h"
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

static std::vector<DocumentTabState> g_documents;
static int g_activeDocument = -1;
static bool g_switchingDocument = false;
static std::vector<DocumentTabState> g_closedDocuments;
static constexpr size_t kMaxClosedDocuments = 20;
static constexpr size_t kTabMemoryCompactThresholdBytes = 256 * 1024;
static bool g_updatingTabs = false;
static bool g_draggingTab = false;
static int g_dragTabIndex = -1;
static int g_hoverTabIndex = -1;
static bool g_hoverTabClose = false;
static bool g_trackingTabsMouse = false;
static bool g_tabsCustomDrawObserved = false;
static constexpr DWORD kSessionMagic = 0x4C4E5331; // "LNS1"
static constexpr DWORD kSessionVersion = 1;
static constexpr DWORD kSessionMaxDocuments = 64;
static constexpr DWORD kSessionMaxStringChars = 8 * 1024 * 1024;
static constexpr DWORD kSessionMaxFileBytes = 64 * 1024 * 1024;
static constexpr UINT_PTR kSessionAutosaveTimerId = 0x4C4E01;
static constexpr UINT kSessionAutosaveIntervalMs = 1500;
static constexpr DWORD kSessionRetryBackoffMs = 10000;
static constexpr bool kEnableCommandBar = false;
static bool g_sessionDirty = false;
static bool g_sessionPersisting = false;
static DWORD g_sessionRetryAtTick = 0;
static HFONT g_hChromeUiFont = nullptr;
static std::vector<std::wstring> g_loadedPrivateFonts;
static HDC g_tabBackbufferDc = nullptr;
static HBITMAP g_tabBackbufferBitmap = nullptr;
static HBITMAP g_tabBackbufferPrevBitmap = nullptr;
static int g_tabBackbufferWidth = 0;
static int g_tabBackbufferHeight = 0;

static std::wstring RuntimeDirectoryPath()
{
    wchar_t modulePath[MAX_PATH] = {};
    if (GetModuleFileNameW(nullptr, modulePath, MAX_PATH) == 0)
        return L".";
    PathRemoveFileSpecW(modulePath);
    return modulePath;
}

static bool ParseEnvBool(const wchar_t *value, bool &out)
{
    if (!value || value[0] == L'\0')
        return false;

    if (lstrcmpiW(value, L"1") == 0 || lstrcmpiW(value, L"true") == 0 || lstrcmpiW(value, L"on") == 0 || lstrcmpiW(value, L"yes") == 0)
    {
        out = true;
        return true;
    }
    if (lstrcmpiW(value, L"0") == 0 || lstrcmpiW(value, L"false") == 0 || lstrcmpiW(value, L"off") == 0 || lstrcmpiW(value, L"no") == 0)
    {
        out = false;
        return true;
    }
    return false;
}

static void ApplyRuntimeFeatureOverrides()
{
    wchar_t envValue[32] = {};
    const DWORD len = GetEnvironmentVariableW(L"TSN_PREMIUM_HEADER", envValue, static_cast<DWORD>(std::size(envValue)));
    if (len == 0 || len >= std::size(envValue))
        return;

    bool parsed = false;
    if (ParseEnvBool(envValue, parsed))
        g_state.premiumHeaderEnabled = parsed;
}

static void LoadBundledFonts()
{
    if (!g_loadedPrivateFonts.empty())
        return;

    static constexpr const wchar_t *kBundledFontFiles[] = {
        L"AkkuratMonoLL-Regular.ttf",
        L"AkkuratMonoLL-Bold.ttf",
        L"AkkuratMonoLL-Italic.ttf",
        L"AkkuratMonoLL-BoldItalic.ttf",
        L"Berkeley Mono Condensed.ttf",
        L"Berkeley Mono Condensed Bold.ttf",
        L"Berkeley Mono Condensed Oblique.ttf",
        L"Berkeley Mono Condensed Bold Oblique.ttf",
    };

    const std::wstring fontDir = RuntimeDirectoryPath() + L"\\fonts\\";
    for (const wchar_t *fileName : kBundledFontFiles)
    {
        std::wstring fontPath = fontDir + fileName;
        if (GetFileAttributesW(fontPath.c_str()) == INVALID_FILE_ATTRIBUTES)
            continue;

        const int added = AddFontResourceExW(fontPath.c_str(), 0, nullptr);
        if (added > 0)
            g_loadedPrivateFonts.push_back(std::move(fontPath));
    }
}

static void UnloadBundledFonts()
{
    for (auto it = g_loadedPrivateFonts.rbegin(); it != g_loadedPrivateFonts.rend(); ++it)
        RemoveFontResourceExW(it->c_str(), 0, nullptr);
    g_loadedPrivateFonts.clear();
}

static HFONT BuildChromeUiFont()
{
    int dpiY = 96;
    HWND ref = g_hwndMain ? g_hwndMain : g_hwndStatus;
    if (ref)
    {
        HDC hdc = GetDC(ref);
        if (hdc)
        {
            const int readDpi = GetDeviceCaps(hdc, LOGPIXELSY);
            if (readDpi > 0)
                dpiY = readDpi;
            ReleaseDC(ref, hdc);
        }
    }

    LOGFONTW lf{};
    lf.lfHeight = -MulDiv(DesignSystem::kChromeFontPointSize, dpiY, 72);
    lf.lfWeight = FW_NORMAL;
    lf.lfQuality = CLEARTYPE_QUALITY;
    wcscpy_s(lf.lfFaceName, DesignSystem::kUiFontPrimary);

    HFONT font = CreateFontIndirectW(&lf);
    if (font)
        return font;
    wcscpy_s(lf.lfFaceName, DesignSystem::kUiFontFallback);
    font = CreateFontIndirectW(&lf);
    if (font)
        return font;
    return TabGetRegularFont();
}

static HFONT ChromeUiFontOrDefault()
{
    return g_hChromeUiFont ? g_hChromeUiFont : TabGetRegularFont();
}

static void RefreshChromeUiFont()
{
    if (g_hChromeUiFont)
    {
        DeleteObject(g_hChromeUiFont);
        g_hChromeUiFont = nullptr;
    }

    g_hChromeUiFont = BuildChromeUiFont();
    if (g_hwndStatus)
        SendMessageW(g_hwndStatus, WM_SETFONT, reinterpret_cast<WPARAM>(ChromeUiFontOrDefault()), TRUE);
    if (g_hwndCommandBar)
        SendMessageW(g_hwndCommandBar, WM_SETFONT, reinterpret_cast<WPARAM>(ChromeUiFontOrDefault()), TRUE);
}

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

static constexpr UINT_PTR kCommandBarButtonIds[] = {
    IDM_FILE_NEW,
    IDM_FILE_OPEN,
    IDM_FILE_SAVE,
    IDM_EDIT_FIND,
    IDM_FORMAT_WORDWRAP,
    IDM_VIEW_DARKMODE,
};
static constexpr size_t kCommandBarButtonCount = sizeof(kCommandBarButtonIds) / sizeof(kCommandBarButtonIds[0]);
static std::wstring g_commandBarLabels[kCommandBarButtonCount];

static std::wstring CommandBarLabelForId(UINT_PTR id)
{
    const auto &lang = GetLangStrings();
    switch (id)
    {
    case IDM_FILE_NEW:
        return MenuLabelForContext(lang.menuNew);
    case IDM_FILE_OPEN:
        return MenuLabelForContext(lang.menuOpen);
    case IDM_FILE_SAVE:
        return MenuLabelForContext(lang.menuSave);
    case IDM_EDIT_FIND:
        return MenuLabelForContext(lang.menuFind);
    case IDM_FORMAT_WORDWRAP:
        return MenuLabelForContext(lang.menuWordWrap);
    case IDM_VIEW_DARKMODE:
        return MenuLabelForContext(lang.menuDarkMode);
    default:
        return L"";
    }
}

static void RefreshCommandBarLabels()
{
    if (!g_hwndCommandBar)
        return;

    const int existingCount = static_cast<int>(SendMessageW(g_hwndCommandBar, TB_BUTTONCOUNT, 0, 0));
    for (int i = existingCount - 1; i >= 0; --i)
        SendMessageW(g_hwndCommandBar, TB_DELETEBUTTON, static_cast<WPARAM>(i), 0);

    for (size_t i = 0; i < kCommandBarButtonCount; ++i)
    {
        std::wstring label = CommandBarLabelForId(kCommandBarButtonIds[i]);
        if (label.empty())
            label = L" ";
        g_commandBarLabels[i] = std::move(label);
    }
    size_t labelIndex = 0;

    std::vector<TBBUTTON> buttons;
    buttons.reserve(kCommandBarButtonCount);

    auto pushButton = [&](UINT_PTR id, BYTE extraStyle = 0)
    {
        TBBUTTON btn{};
        btn.iBitmap = I_IMAGENONE;
        btn.idCommand = static_cast<int>(id);
        btn.fsState = TBSTATE_ENABLED;
        btn.fsStyle = static_cast<BYTE>(BTNS_AUTOSIZE | BTNS_SHOWTEXT | BTNS_BUTTON | extraStyle);
        if (labelIndex < kCommandBarButtonCount)
            btn.iString = reinterpret_cast<INT_PTR>(g_commandBarLabels[labelIndex].c_str());
        else
            btn.iString = reinterpret_cast<INT_PTR>(L" ");
        ++labelIndex;
        buttons.push_back(btn);
    };

    pushButton(IDM_FILE_NEW);
    pushButton(IDM_FILE_OPEN);
    pushButton(IDM_FILE_SAVE);
    pushButton(IDM_EDIT_FIND);
    pushButton(IDM_FORMAT_WORDWRAP, BTNS_CHECK);
    pushButton(IDM_VIEW_DARKMODE, BTNS_CHECK);

    SendMessageW(g_hwndCommandBar, TB_ADDBUTTONSW, static_cast<WPARAM>(buttons.size()), reinterpret_cast<LPARAM>(buttons.data()));
    SendMessageW(g_hwndCommandBar, TB_AUTOSIZE, 0, 0);
}

static void RefreshCommandBarMetrics()
{
    if (!g_hwndCommandBar)
        return;

    SendMessageW(g_hwndCommandBar, WM_SETFONT, reinterpret_cast<WPARAM>(ChromeUiFontOrDefault()), TRUE);
    SendMessageW(g_hwndCommandBar,
                 TB_SETPADDING,
                 0,
                 MAKELPARAM(TabScalePx(DesignSystem::kCommandBarPaddingHPx), TabScalePx(DesignSystem::kCommandBarPaddingVPx)));
    SendMessageW(g_hwndCommandBar, TB_SETINDENT, TabScalePx(DesignSystem::kCommandBarIndentPx), 0);
    SendMessageW(g_hwndCommandBar, TB_AUTOSIZE, 0, 0);
}

static void RefreshCommandBarTheme()
{
    if (!g_hwndCommandBar)
        return;

    InvalidateRect(g_hwndCommandBar, nullptr, FALSE);
}

static void InitializeCommandBar()
{
    if (!g_hwndCommandBar)
        return;

    SendMessageW(g_hwndCommandBar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
    SendMessageW(g_hwndCommandBar, TB_SETMAXTEXTROWS, 1, 0);
    SendMessageW(g_hwndCommandBar,
                 TB_SETEXTENDEDSTYLE,
                 0,
                 TBSTYLE_EX_MIXEDBUTTONS | TBSTYLE_EX_DOUBLEBUFFER | TBSTYLE_EX_HIDECLIPPEDBUTTONS);
    RefreshCommandBarLabels();
    RefreshCommandBarMetrics();
    RefreshCommandBarTheme();
}

static void RefreshCommandBarStates()
{
    if (!g_hwndCommandBar)
        return;

    const BOOL wrapEnabled = (g_state.wordWrap && !g_state.largeFileMode) ? TRUE : FALSE;
    const BOOL darkEnabled = IsDarkMode() ? TRUE : FALSE;
    const BOOL wrapToggleEnabled = g_state.largeFileMode ? FALSE : TRUE;

    SendMessageW(g_hwndCommandBar, TB_ENABLEBUTTON, IDM_FORMAT_WORDWRAP, MAKELONG(wrapToggleEnabled, 0));
    SendMessageW(g_hwndCommandBar, TB_CHECKBUTTON, IDM_FORMAT_WORDWRAP, MAKELONG(wrapEnabled, 0));
    SendMessageW(g_hwndCommandBar, TB_CHECKBUTTON, IDM_VIEW_DARKMODE, MAKELONG(darkEnabled, 0));
}

static void UpdateRuntimeMenuStates();
static void RebuildTabsControl();
static void LoadStateFromDocument(int index);
static bool OpenPathInTabs(const std::wstring &path, bool forceReplaceCurrent = false);
static void ResetActiveDocumentToUntitled();
static bool ConfirmCloseForCurrentStartupBehavior();

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

static void EnsureSingleDocumentModel(bool captureEditorText)
{
    DocumentTabState doc;
    if (captureEditorText || g_documents.empty() || g_activeDocument < 0 || g_activeDocument >= static_cast<int>(g_documents.size()))
    {
        doc.text = GetEditorText();
        if (!g_state.largeFileMode)
        {
            doc.richText = GetEditorRichText();
            doc.hasRichText = !doc.richText.empty();
        }
    }
    else
    {
        doc.text = g_documents[g_activeDocument].text;
        doc.richText = g_documents[g_activeDocument].richText;
        doc.hasRichText = g_documents[g_activeDocument].hasRichText;
    }

    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    doc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateDocumentTextBytes(doc.text);
    doc.largeFileMode = g_state.largeFileMode || ShouldUseLargeDocumentMode(doc.sourceBytes);

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

static std::wstring DocumentTabLabel(const DocumentTabState &doc)
{
    const auto &lang = GetLangStrings();
    std::wstring label = doc.filePath.empty() ? lang.untitled : PathFindFileNameW(doc.filePath.c_str());
    if (doc.modified)
        label.insert(label.begin(), L'*');
    return label;
}

static bool TryGetTabInvalidateRect(HWND hwnd, int index, RECT &rc)
{
    if (!hwnd || index < 0)
        return false;

    RECT itemRect{};
    if (!TabCtrl_GetItemRect(hwnd, index, &itemRect))
        return false;

    itemRect.left = std::max<LONG>(0, itemRect.left - TabScalePx(1));
    itemRect.right += TabScalePx(1);
    itemRect.top = std::max<LONG>(0, itemRect.top);
    itemRect.bottom += TabScalePx(1);
    rc = itemRect;
    return true;
}

static void InvalidateTabItem(HWND hwnd, int index)
{
    if (index < 0)
        return;

    RECT rc{};
    if (TryGetTabInvalidateRect(hwnd, index, rc))
    {
        InvalidateRect(hwnd, &rc, FALSE);
        return;
    }
    InvalidateRect(hwnd, nullptr, FALSE);
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
    // Ensure visual state of active/non-active tabs is refreshed in one pass.
    RedrawWindow(g_hwndTabs, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
}

static bool IsTabCloseRectHit(int index, POINT ptClient)
{
    (void)index;
    (void)ptClient;
    return false;
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
    const bool tabLabelChanged = (doc.filePath != g_state.filePath) || (doc.modified != g_state.modified);
    if (includeText)
    {
        doc.text = GetEditorText();
        if (!g_state.largeFileMode)
        {
            doc.richText = GetEditorRichText();
            doc.hasRichText = !doc.richText.empty();
        }
        else
        {
            doc.richText.clear();
            doc.hasRichText = false;
        }
        doc.needsReloadFromDisk = false;
    }
    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    doc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateDocumentTextBytes(doc.text);
    doc.largeFileMode = g_state.largeFileMode || ShouldUseLargeDocumentMode(doc.sourceBytes);
    if (tabLabelChanged)
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
        if (!SessionLoadDocumentTextFromDisk(doc))
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
    if (doc.hasRichText && !doc.richText.empty() && !doc.largeFileMode)
        SetEditorRichText(doc.richText);
    else
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

    const int previousIndex = g_activeDocument;
    if (previousIndex >= 0)
        SyncDocumentFromState(previousIndex, true);

    g_activeDocument = index;
    if (previousIndex >= 0)
        TabCompactDocumentTextIfEligible(g_documents, previousIndex, g_activeDocument, kTabMemoryCompactThresholdBytes, SessionPathExists);
    if (g_hwndTabs)
        TabCtrl_SetCurSel(g_hwndTabs, index);
    LoadStateFromDocument(index);
    if (g_hwndTabs)
    {
        InvalidateTabItem(g_hwndTabs, previousIndex);
        InvalidateTabItem(g_hwndTabs, index);
        // Force full strip repaint so previous active tab doesn't retain stale style.
        RedrawWindow(g_hwndTabs, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
    }
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
    if (!g_state.largeFileMode)
    {
        doc.richText = GetEditorRichText();
        doc.hasRichText = !doc.richText.empty();
    }
    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    doc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateDocumentTextBytes(doc.text);
    doc.largeFileMode = g_state.largeFileMode || ShouldUseLargeDocumentMode(doc.sourceBytes);

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
    const int previousIndex = g_activeDocument;
    if (previousIndex >= 0)
        SyncDocumentFromState(previousIndex, true);

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
    if (previousIndex >= 0)
        TabCompactDocumentTextIfEligible(g_documents, previousIndex, g_activeDocument, kTabMemoryCompactThresholdBytes, SessionPathExists);
    LoadStateFromDocument(index);
    UpdateRuntimeMenuStates();
    MarkSessionDirty();
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

    const int existingIndex = TabFindDocumentByPath(g_documents, path, SessionNormalizePathForCompare);
    if (existingIndex >= 0)
    {
        SwitchToDocument(existingIndex);
        return true;
    }

    const int previousIndex = g_activeDocument;

    bool createdNewTab = false;
    if (!(g_activeDocument >= 0 &&
          g_activeDocument < static_cast<int>(g_documents.size()) &&
          TabIsEmptyUntitled(g_documents[g_activeDocument])))
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
        if (!g_state.largeFileMode)
        {
            activeDoc.richText = GetEditorRichText();
            activeDoc.hasRichText = !activeDoc.richText.empty();
        }
        else
        {
            activeDoc.richText.clear();
            activeDoc.hasRichText = false;
        }
        activeDoc.filePath = g_state.filePath;
        activeDoc.modified = g_state.modified;
        activeDoc.encoding = g_state.encoding;
        activeDoc.lineEnding = g_state.lineEnding;
        activeDoc.sourceBytes = (g_state.largeFileBytes > 0) ? g_state.largeFileBytes : EstimateDocumentTextBytes(activeDoc.text);
        activeDoc.largeFileMode = g_state.largeFileMode || ShouldUseLargeDocumentMode(activeDoc.sourceBytes);
    }

    const bool resumeAll = (g_state.startupBehavior == StartupBehavior::ResumeAll);
    const bool startupClassic = (g_state.startupBehavior == StartupBehavior::Classic);
    const std::wstring sessionFilePath = SessionRuntimeFilePath();

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
        saveOk = SessionWriteSnapshot(sessionFilePath,
                                      g_documents,
                                      g_activeDocument,
                                      kSessionMagic,
                                      kSessionVersion,
                                      kSessionMaxDocuments,
                                      kSessionMaxStringChars,
                                      kSessionMaxFileBytes);
    }
    else
    {
        DeleteFileW(sessionFilePath.c_str());
    }

    if (persistPathFallback)
    {
        std::vector<std::wstring> sessionPaths;
        int activePathIndex = -1;
        TabBuildPathSessionFallback(g_documents, g_activeDocument, sessionPaths, activePathIndex, SessionNormalizePathForCompare);
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
        const std::wstring sessionFilePath = SessionRuntimeFilePath();
        TabSessionSnapshot snapshot;
        if (SessionReadSnapshot(sessionFilePath,
                                snapshot,
                                kSessionMagic,
                                kSessionVersion,
                                kSessionMaxDocuments,
                                kSessionMaxStringChars,
                                kSessionMaxFileBytes,
                                true))
        {
            g_documents = std::move(snapshot.documents);
            g_closedDocuments.clear();
            g_activeDocument = snapshot.activeDocument;
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
        if (path.empty() || !SessionPathExists(path))
            return false;
        return OpenPathInTabs(path);
    }

    bool openedAny = false;
    for (const auto &path : sessionPaths)
    {
        if (path.empty() || !SessionPathExists(path))
            continue;
        if (OpenPathInTabs(path))
            openedAny = true;
    }

    if (!openedAny)
        return false;

    if (activePathIndex < 0 || activePathIndex >= static_cast<int>(sessionPaths.size()))
        return true;

    const int activeDocIndex = TabFindDocumentByPath(g_documents, sessionPaths[activePathIndex], SessionNormalizePathForCompare);
    if (activeDocIndex >= 0)
        SwitchToDocument(activeDocIndex);
    UpdateRuntimeMenuStates();
    return true;
}

static void OpenFileInNewDocumentTabDialog()
{
    const auto &lang = GetLangStrings();
    const std::wstring filter = BuildTextDocumentsFilter(lang);
    wchar_t path[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndMain;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_ENABLESIZING;
    if (GetOpenFileNameW(&ofn))
        OpenPathInTabs(path);
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
        TabPushClosedDocument(g_closedDocuments, g_documents[g_activeDocument], kMaxClosedDocuments);
        ResetActiveDocumentToUntitled();
        SyncDocumentFromState(g_activeDocument, true);
        UpdateRuntimeMenuStates();
        MarkSessionDirty();
        return;
    }

    if (!ConfirmDiscard())
        return;

    const int closingIndex = index;
    TabPushClosedDocument(g_closedDocuments, g_documents[closingIndex], kMaxClosedDocuments);
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

    const int previousIndex = g_activeDocument;
    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    DocumentTabState doc = g_closedDocuments.back();
    g_closedDocuments.pop_back();
    g_documents.push_back(doc);
    g_activeDocument = static_cast<int>(g_documents.size()) - 1;
    if (previousIndex >= 0)
        TabCompactDocumentTextIfEligible(g_documents, previousIndex, g_activeDocument, kTabMemoryCompactThresholdBytes, SessionPathExists);
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

    DocumentTabState moving = std::move(g_documents[fromIndex]);
    g_documents.erase(g_documents.begin() + fromIndex);
    g_documents.insert(g_documents.begin() + toIndex, std::move(moving));

    if (g_activeDocument == fromIndex)
    {
        g_activeDocument = toIndex;
    }
    else if (fromIndex < toIndex)
    {
        if (g_activeDocument > fromIndex && g_activeDocument <= toIndex)
            --g_activeDocument;
    }
    else
    {
        if (g_activeDocument >= toIndex && g_activeDocument < fromIndex)
            ++g_activeDocument;
    }

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
    EnableMenuItem(hMenu, IDM_FORMAT_BOLD, MF_BYCOMMAND | (g_state.largeFileMode ? MF_GRAYED : MF_ENABLED));
    EnableMenuItem(hMenu, IDM_FORMAT_ITALIC, MF_BYCOMMAND | (g_state.largeFileMode ? MF_GRAYED : MF_ENABLED));
    EnableMenuItem(hMenu, IDM_FORMAT_STRIKETHROUGH, MF_BYCOMMAND | (g_state.largeFileMode ? MF_GRAYED : MF_ENABLED));

    const bool tabModeActive = g_state.useTabs;
    const bool hasMultipleTabs = g_documents.size() > 1;
    const bool hasClosedTabs = !g_closedDocuments.empty();

    EnableMenuItem(hMenu, IDM_FILE_CLOSETAB, MF_BYCOMMAND | ((tabModeActive && hasMultipleTabs) ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(hMenu, IDM_FILE_REOPENCLOSEDTAB, MF_BYCOMMAND | ((tabModeActive && hasClosedTabs) ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(hMenu, IDM_FILE_NEXTTAB, MF_BYCOMMAND | ((tabModeActive && hasMultipleTabs) ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(hMenu, IDM_FILE_PREVTAB, MF_BYCOMMAND | ((tabModeActive && hasMultipleTabs) ? MF_ENABLED : MF_GRAYED));

    RefreshCommandBarStates();

    DrawMenuBar(g_hwndMain);
}

struct CommandBarPalette
{
    COLORREF background;
    COLORREF border;
    COLORREF text;
    COLORREF disabledText;
    COLORREF hoverBg;
    COLORREF checkedBg;
};

static CommandBarPalette GetCommandBarPalette(bool dark)
{
    if (dark)
    {
        return {
            ThemeColorMenuBackground(true),
            ThemeColorChromeBorder(true),
            ThemeColorMenuText(true),
            ThemeColorMenuDisabledText(true),
            ThemeColorMenuHoverBackground(true),
            ThemeColorMenuHoverBackground(true),
        };
    }

    return {
        ThemeColorMenuBackground(false),
        ThemeColorChromeBorder(false),
        ThemeColorMenuText(false),
        ThemeColorMenuDisabledText(false),
        ThemeColorMenuHoverBackground(false),
        ThemeColorMenuHoverBackground(false),
    };
}

static LRESULT HandleCommandBarCustomDraw(LPNMTBCUSTOMDRAW draw)
{
    if (!draw || !g_hwndCommandBar)
        return CDRF_DODEFAULT;

    const bool dark = IsDarkMode();
    const CommandBarPalette palette = GetCommandBarPalette(dark);

    if (draw->nmcd.dwDrawStage == CDDS_PREPAINT)
    {
        RECT rc{};
        GetClientRect(g_hwndCommandBar, &rc);
        HBRUSH hbrBg = CreateSolidBrush(palette.background);
        if (hbrBg)
        {
            FillRect(draw->nmcd.hdc, &rc, hbrBg);
            DeleteObject(hbrBg);
        }

        RECT border = rc;
        border.top = std::max(border.top, border.bottom - std::max(1, TabScalePx(DesignSystem::kChromeStrokePx)));
        HBRUSH hbrBorder = CreateSolidBrush(palette.border);
        if (hbrBorder)
        {
            FillRect(draw->nmcd.hdc, &border, hbrBorder);
            DeleteObject(hbrBorder);
        }
        return CDRF_NOTIFYITEMDRAW;
    }

    if (draw->nmcd.dwDrawStage != CDDS_ITEMPREPAINT)
        return CDRF_DODEFAULT;

    const bool disabled = (draw->nmcd.uItemState & CDIS_DISABLED) != 0;
    const bool checked = (draw->nmcd.uItemState & CDIS_CHECKED) != 0;
    const bool hot = (draw->nmcd.uItemState & (CDIS_HOT | CDIS_SELECTED)) != 0;

    RECT rc = draw->nmcd.rc;
    const int hoverInset = std::max(0, TabScalePx(DesignSystem::kCommandBarHoverInsetPx));
    if (hoverInset > 0)
        InflateRect(&rc, -hoverInset, -hoverInset);
    if (checked || hot)
    {
        HBRUSH hbrItem = CreateSolidBrush(checked ? palette.checkedBg : palette.hoverBg);
        if (hbrItem)
        {
            FillRect(draw->nmcd.hdc, &rc, hbrItem);
            DeleteObject(hbrItem);
        }
    }

    const int itemIndex = static_cast<int>(SendMessageW(g_hwndCommandBar, TB_COMMANDTOINDEX, draw->nmcd.dwItemSpec, 0));
    const int itemCount = static_cast<int>(SendMessageW(g_hwndCommandBar, TB_BUTTONCOUNT, 0, 0));
    if (itemIndex >= 0 && itemIndex < itemCount - 1)
    {
        RECT separator = draw->nmcd.rc;
        const int stroke = std::max(1, TabScalePx(DesignSystem::kChromeStrokePx));
        separator.left = std::max(separator.left, separator.right - stroke);
        separator.top += stroke;
        separator.bottom -= stroke;
        if (separator.bottom <= separator.top)
            separator.bottom = separator.top + 1;
        HBRUSH hbrSeparator = CreateSolidBrush(palette.border);
        if (hbrSeparator)
        {
            FillRect(draw->nmcd.hdc, &separator, hbrSeparator);
            DeleteObject(hbrSeparator);
        }
    }

    draw->clrText = disabled ? palette.disabledText : palette.text;
    SetBkMode(draw->nmcd.hdc, TRANSPARENT);
    SelectObject(draw->nmcd.hdc, ChromeUiFontOrDefault());
    return static_cast<LRESULT>(TBCDRF_USECDCOLORS | TBCDRF_NOEDGES | TBCDRF_NOOFFSET | TBCDRF_NOBACKGROUND | CDRF_NEWFONT);
}

// TabPaintPalette and GetTabPaintPalette moved to theme.h/cpp

static void FillSolidRectDc(HDC hdc, const RECT &rc, COLORREF color)
{
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(DC_BRUSH));
    const COLORREF oldBrushColor = SetDCBrushColor(hdc, color);
    FillRect(hdc, &rc, reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
    SetDCBrushColor(hdc, oldBrushColor);
    SelectObject(hdc, oldBrush);
}

static COLORREF BlendColor(COLORREF fg, COLORREF bg, int fgPercent)
{
    fgPercent = std::clamp(fgPercent, 0, 100);
    const int bgPercent = 100 - fgPercent;
    const int r = (GetRValue(fg) * fgPercent + GetRValue(bg) * bgPercent) / 100;
    const int g = (GetGValue(fg) * fgPercent + GetGValue(bg) * bgPercent) / 100;
    const int b = (GetBValue(fg) * fgPercent + GetBValue(bg) * bgPercent) / 100;
    return RGB(r, g, b);
}

static RECT MenuTextRectWithInsets(const RECT &src)
{
    RECT rc = src;
    const int hPad = std::max(0, TabScalePx(DesignSystem::kMenuTextPaddingHPx));
    const int vPad = TabScalePx(DesignSystem::kMenuTextPaddingVPx);
    const int width = rc.right - rc.left;
    if (width > (hPad * 2 + 2))
    {
        rc.left += hPad;
        rc.right -= hPad;
    }
    rc.top += vPad;
    rc.bottom -= vPad;
    if (rc.right <= rc.left)
        rc.right = rc.left + 1;
    if (rc.bottom <= rc.top)
        rc.bottom = rc.top + 1;
    return rc;
}

static std::wstring GetMenuItemLabelByCommandId(UINT commandId)
{
    if (!g_hwndMain || commandId == 0)
        return L"";

    HMENU hMenu = GetMenu(g_hwndMain);
    if (!hMenu)
        return L"";

    std::wstring text(256, L'\0');
    int copied = GetMenuStringW(hMenu, commandId, text.data(), static_cast<int>(text.size()), MF_BYCOMMAND);
    if (copied <= 0)
        return L"";

    text.resize(static_cast<size_t>(copied));
    return text;
}

static void SplitMenuLabelAndShortcut(const std::wstring &fullText, std::wstring &label, std::wstring &shortcut)
{
    label = fullText;
    shortcut.clear();

    const size_t tabPos = fullText.find(L'\t');
    if (tabPos == std::wstring::npos)
        return;

    label = fullText.substr(0, tabPos);
    shortcut = fullText.substr(tabPos + 1);
}

static bool HandlePopupMenuMeasureItem(MEASUREITEMSTRUCT *measure)
{
    if (!measure || measure->CtlType != ODT_MENU)
        return false;

    const std::wstring fullText = GetMenuItemLabelByCommandId(static_cast<UINT>(measure->itemID));
    if (fullText.empty())
        return false;

    std::wstring label;
    std::wstring shortcut;
    SplitMenuLabelAndShortcut(fullText, label, shortcut);

    HDC hdc = GetDC(g_hwndMain ? g_hwndMain : GetDesktopWindow());
    if (!hdc)
        return false;

    HFONT hFont = ChromeUiFontOrDefault();
    HGDIOBJ oldFont = SelectObject(hdc, hFont);

    SIZE labelSize{};
    GetTextExtentPoint32W(hdc, label.c_str(), static_cast<int>(label.size()), &labelSize);

    SIZE shortcutSize{};
    if (!shortcut.empty())
        GetTextExtentPoint32W(hdc, shortcut.c_str(), static_cast<int>(shortcut.size()), &shortcutSize);

    TEXTMETRICW tm{};
    GetTextMetricsW(hdc, &tm);

    if (oldFont)
        SelectObject(hdc, oldFont);
    ReleaseDC(g_hwndMain ? g_hwndMain : GetDesktopWindow(), hdc);

    const UINT leftInset = static_cast<UINT>(TabScalePx(14));
    const UINT rightInset = static_cast<UINT>(TabScalePx(14));
    const UINT checkSlot = static_cast<UINT>(TabScalePx(14));
    const UINT accelGap = shortcut.empty() ? 0u : static_cast<UINT>(TabScalePx(20));
    const UINT textGap = static_cast<UINT>(TabScalePx(8));
    const UINT verticalPadding = static_cast<UINT>(TabScalePx(5));

    measure->itemWidth = leftInset + checkSlot + textGap + static_cast<UINT>(labelSize.cx) + accelGap + static_cast<UINT>(shortcutSize.cx) + rightInset;
    measure->itemHeight = std::max<UINT>(static_cast<UINT>(tm.tmHeight) + verticalPadding * 2, static_cast<UINT>(TabScalePx(26)));
    return true;
}

static bool HandlePopupMenuDrawItem(DRAWITEMSTRUCT *draw)
{
    if (!draw || draw->CtlType != ODT_MENU)
        return false;

    const std::wstring fullText = GetMenuItemLabelByCommandId(static_cast<UINT>(draw->itemID));
    if (fullText.empty())
        return false;

    std::wstring label;
    std::wstring shortcut;
    SplitMenuLabelAndShortcut(fullText, label, shortcut);

    const bool dark = IsDarkMode();
    const bool disabled = (draw->itemState & ODS_DISABLED) != 0;
    const bool selected = (draw->itemState & ODS_SELECTED) != 0;
    const bool checked = (draw->itemState & ODS_CHECKED) != 0;

    const COLORREF bg = selected ? ThemeColorMenuHoverBackground(dark) : ThemeColorMenuBackground(dark);
    const COLORREF text = ThemeColorMenuText(dark);
    const COLORREF disabledText = BlendColor(text, bg, dark ? 45 : 35);

    FillSolidRectDc(draw->hDC, draw->rcItem, bg);
    SetBkMode(draw->hDC, TRANSPARENT);

    HFONT hFont = ChromeUiFontOrDefault();
    HGDIOBJ oldFont = SelectObject(draw->hDC, hFont);

    RECT rc = draw->rcItem;
    const int leftInset = TabScalePx(14);
    const int rightInset = TabScalePx(14);
    const int checkSlot = TabScalePx(14);
    const int textGap = TabScalePx(8);
    const int accelGap = shortcut.empty() ? 0 : TabScalePx(20);
    int shortcutWidth = 0;
    if (!shortcut.empty())
    {
        SIZE shortcutSize{};
        if (GetTextExtentPoint32W(draw->hDC, shortcut.c_str(), static_cast<int>(shortcut.size()), &shortcutSize))
            shortcutWidth = shortcutSize.cx;
    }

    RECT rcLabel = rc;
    rcLabel.left += leftInset + checkSlot + textGap;
    rcLabel.right -= (rightInset + accelGap + shortcutWidth);

    RECT rcShortcut = rc;
    rcShortcut.left = std::max(rcShortcut.left, rcLabel.right + textGap);
    rcShortcut.right -= rightInset;

    SetTextColor(draw->hDC, disabled ? disabledText : text);
    DrawTextW(draw->hDC, label.c_str(), -1, &rcLabel, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (!shortcut.empty())
        DrawTextW(draw->hDC, shortcut.c_str(), -1, &rcShortcut, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if (checked)
    {
        RECT rcCheck = rc;
        rcCheck.left += leftInset;
        rcCheck.right = rcCheck.left + checkSlot;
        DrawTextW(draw->hDC, L"\u2713", -1, &rcCheck, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    if (oldFont)
        SelectObject(draw->hDC, oldFont);
    return true;
}

[[maybe_unused]] static void DrawRectOutlineDc(HDC hdc, const RECT &rc, COLORREF color, bool includeBottom = true)
{
    const int stroke = std::max(1, TabScalePx(1));

    RECT top = rc;
    top.bottom = std::min(top.bottom, top.top + stroke);
    FillSolidRectDc(hdc, top, color);

    RECT left = rc;
    left.right = std::min(left.right, left.left + stroke);
    FillSolidRectDc(hdc, left, color);

    RECT right = rc;
    right.left = std::max(right.left, right.right - stroke);
    FillSolidRectDc(hdc, right, color);

    if (includeBottom)
    {
        RECT bottom = rc;
        bottom.top = std::max(bottom.top, bottom.bottom - stroke);
        FillSolidRectDc(hdc, bottom, color);
    }
}

static void DrawTabStripBackground(HDC hdc, const RECT &rcClient, const TabPaintPalette &palette)
{
    FillSolidRectDc(hdc, rcClient, palette.stripBg);
    RECT bottomDivider = rcClient;
    const int stroke = std::max(1, TabScalePx(DesignSystem::kChromeStrokePx));
    bottomDivider.top = std::max(bottomDivider.top, bottomDivider.bottom - stroke);
    FillSolidRectDc(hdc, bottomDivider, palette.stripBorder);
}

static void DrawTabItemVisual(HDC hdc, int index, const RECT &rawItemRect, const TabPaintPalette &palette)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    const int visualActiveIndex = g_hwndTabs ? TabCtrl_GetCurSel(g_hwndTabs) : g_activeDocument;
    const bool selected = (index == visualActiveIndex);
    const bool hovered = (index == g_hoverTabIndex);
    const bool dark = IsDarkMode();
    const int tabCount = std::min(TabCtrl_GetItemCount(g_hwndTabs), static_cast<int>(g_documents.size()));

    RECT tabsClient{};
    GetClientRect(g_hwndTabs, &tabsClient);

    RECT contentRect = rawItemRect;
    if (index == 0)
    {
        contentRect.left = 0;
    }
    else
    {
        RECT prevRect{};
        if (TabCtrl_GetItemRect(g_hwndTabs, index - 1, &prevRect))
            contentRect.left = prevRect.right;
    }

    if (index == tabCount - 1)
    {
        contentRect.right = tabsClient.right;
    }
    else
    {
        RECT nextRect{};
        if (TabCtrl_GetItemRect(g_hwndTabs, index + 1, &nextRect))
            contentRect.right = nextRect.left;
    }

    contentRect.top = 0;
    contentRect.bottom = tabsClient.bottom;
    if (contentRect.right <= contentRect.left)
        contentRect.right = contentRect.left + 1;
    if (contentRect.bottom <= contentRect.top)
        contentRect.bottom = contentRect.top + 1;

    const COLORREF activeSurface = ThemeColorEditorBackground(dark);
    const int stroke = std::max(1, TabScalePx(DesignSystem::kTabSeamStrokePx));
    const COLORREF selectedBg = palette.activeBg;
    const COLORREF itemBg = selected ? selectedBg : (hovered ? palette.hoverBg : palette.inactiveBg);
    FillSolidRectDc(hdc, contentRect, itemBg);

    if (index > 0)
    {
        RECT separator = contentRect;
        const int inset = TabScalePx(DesignSystem::kTabSeparatorInsetYPx);
        separator.top += inset;
        separator.bottom -= inset;
        if (separator.bottom <= separator.top)
            separator.bottom = separator.top + 1;
        separator.right = std::min(separator.right, separator.left + stroke);
        const COLORREF separatorColor = BlendColor(
            palette.stripBorder,
            itemBg,
            DesignSystem::kTabSeparatorAlphaPct);
        FillSolidRectDc(hdc, separator, separatorColor);
    }

    HFONT drawFont = selected ? TabGetActiveFont() : TabGetRegularFont();
    HGDIOBJ oldFont = nullptr;
    if (drawFont)
        oldFont = SelectObject(hdc, drawFont);

    RECT textRect = contentRect;
    textRect.left += TabScalePx(DesignSystem::kTabTextPaddingHPx);
    textRect.right -= TabScalePx(DesignSystem::kTabTextPaddingHPx);
    if (textRect.right < textRect.left)
        textRect.right = textRect.left;
    std::wstring label = DocumentTabLabel(g_documents[index]);
    SetBkMode(hdc, TRANSPARENT);
    const COLORREF hoverTextColor = BlendColor(
        palette.activeTextColor,
        palette.textColor,
        dark ? 45 : 55);
    const COLORREF tabTextColor = selected
                                      ? palette.activeTextColor
                                      : (hovered ? hoverTextColor : palette.textColor);
    SetTextColor(hdc, tabTextColor);
    DrawTextW(hdc, label.c_str(), -1, &textRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
    if (oldFont)
        SelectObject(hdc, oldFont);

    if (selected)
    {
        // Cut the strip divider under active tab to visually attach tab to editor.
        const int seamStroke = stroke;
        RECT seamRect = contentRect;
        seamRect.top = std::max(seamRect.top, seamRect.bottom - seamStroke);
        FillSolidRectDc(hdc, seamRect, activeSurface);
    }
}

[[maybe_unused]] static void PaintTabStripVisual(HDC hdc)
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
        if (i == g_activeDocument)
            continue;
        RECT itemRect{};
        if (TabCtrl_GetItemRect(g_hwndTabs, i, &itemRect))
            DrawTabItemVisual(hdc, i, itemRect, palette);
    }
    if (g_activeDocument >= 0 && g_activeDocument < count)
    {
        RECT activeRect{};
        if (TabCtrl_GetItemRect(g_hwndTabs, g_activeDocument, &activeRect))
            DrawTabItemVisual(hdc, g_activeDocument, activeRect, palette);
    }
}

static void ReleaseTabBackbuffer()
{
    if (g_tabBackbufferDc && g_tabBackbufferPrevBitmap)
        SelectObject(g_tabBackbufferDc, g_tabBackbufferPrevBitmap);

    if (g_tabBackbufferBitmap)
    {
        DeleteObject(g_tabBackbufferBitmap);
        g_tabBackbufferBitmap = nullptr;
    }
    if (g_tabBackbufferDc)
    {
        DeleteDC(g_tabBackbufferDc);
        g_tabBackbufferDc = nullptr;
    }
    g_tabBackbufferPrevBitmap = nullptr;
    g_tabBackbufferWidth = 0;
    g_tabBackbufferHeight = 0;
}

static bool EnsureTabBackbuffer(HDC targetHdc, int width, int height, HDC &outMemDc)
{
    outMemDc = nullptr;
    if (!targetHdc || width <= 0 || height <= 0)
        return false;

    if (!g_tabBackbufferDc)
    {
        g_tabBackbufferDc = CreateCompatibleDC(targetHdc);
        if (!g_tabBackbufferDc)
            return false;
    }

    const bool sizeChanged = (g_tabBackbufferWidth != width) || (g_tabBackbufferHeight != height);
    if (g_tabBackbufferBitmap && sizeChanged)
    {
        if (g_tabBackbufferPrevBitmap)
            SelectObject(g_tabBackbufferDc, g_tabBackbufferPrevBitmap);
        DeleteObject(g_tabBackbufferBitmap);
        g_tabBackbufferBitmap = nullptr;
    }

    if (!g_tabBackbufferBitmap)
    {
        g_tabBackbufferBitmap = CreateCompatibleBitmap(targetHdc, width, height);
        if (!g_tabBackbufferBitmap)
        {
            ReleaseTabBackbuffer();
            return false;
        }
        g_tabBackbufferWidth = width;
        g_tabBackbufferHeight = height;
    }

    HGDIOBJ oldBitmap = SelectObject(g_tabBackbufferDc, g_tabBackbufferBitmap);
    if (!g_tabBackbufferPrevBitmap && oldBitmap && oldBitmap != HGDI_ERROR)
        g_tabBackbufferPrevBitmap = reinterpret_cast<HBITMAP>(oldBitmap);
    outMemDc = g_tabBackbufferDc;
    return true;
}

static void PaintTabStripBuffered(HWND hwnd, HDC targetHdc)
{
    if (!hwnd || !targetHdc)
        return;

    RECT rcClient{};
    GetClientRect(hwnd, &rcClient);
    const int width = std::max(1, static_cast<int>(rcClient.right - rcClient.left));
    const int height = std::max(1, static_cast<int>(rcClient.bottom - rcClient.top));

    HDC memDc = nullptr;
    if (!EnsureTabBackbuffer(targetHdc, width, height, memDc))
    {
        PaintTabStripVisual(targetHdc);
        return;
    }

    PaintTabStripVisual(memDc);

    RECT blitRect{};
    if (GetClipBox(targetHdc, &blitRect) == ERROR ||
        blitRect.right <= blitRect.left ||
        blitRect.bottom <= blitRect.top)
    {
        blitRect.left = 0;
        blitRect.top = 0;
        blitRect.right = width;
        blitRect.bottom = height;
    }

    const int blitWidth = std::max(1, static_cast<int>(blitRect.right - blitRect.left));
    const int blitHeight = std::max(1, static_cast<int>(blitRect.bottom - blitRect.top));
    BitBlt(targetHdc, blitRect.left, blitRect.top, blitWidth, blitHeight, memDc, blitRect.left, blitRect.top, SRCCOPY);
}

static void UpdateTabsHoverState(HWND hwnd, POINT ptClient)
{
    TCHITTESTINFO hit{};
    hit.pt = ptClient;
    const int prevHoverIndex = g_hoverTabIndex;
    const int hoverIndex = TabCtrl_HitTest(hwnd, &hit);
    if (hoverIndex != g_hoverTabIndex)
    {
        g_hoverTabIndex = hoverIndex;
        g_hoverTabClose = false;
        if (prevHoverIndex != hoverIndex)
        {
            InvalidateTabItem(hwnd, prevHoverIndex);
            InvalidateTabItem(hwnd, hoverIndex);
        }
        else
        {
            InvalidateTabItem(hwnd, hoverIndex);
        }
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

[[maybe_unused]] static LRESULT HandleTabsCustomDraw(LPNMCUSTOMDRAW draw)
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
    RECT itemRect = draw->rc;
    TabCtrl_GetItemRect(g_hwndTabs, index, &itemRect);
    DrawTabItemVisual(draw->hdc, index, itemRect, palette);
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
        ReleaseTabBackbuffer();
        TabRefreshDpi();
        TabRefreshVisualMetrics();
        g_tabsCustomDrawObserved = false;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    case WM_ERASEBKGND:
        return 1;
    case WM_PRINTCLIENT:
    case WM_PRINT:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        PaintTabStripBuffered(hwnd, hdc);
        return 0;
    }
    case WM_PAINT:
    {
        TabSpinAttachIfNeeded(hwnd);
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        PaintTabStripBuffered(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_LBUTTONDOWN:
    {
        if (!g_state.useTabs || g_documents.size() <= 1)
            break;

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

        if (!g_draggingTab)
            break;
        if (!(wParam & MK_LBUTTON))
        {
            g_draggingTab = false;
            g_dragTabIndex = -1;
            if (GetCapture() == hwnd)
                ReleaseCapture();
            break;
        }

        TCHITTESTINFO hit{};
        hit.pt = pt;
        int hoverIndex = TabCtrl_HitTest(hwnd, &hit);
        if (hoverIndex >= 0 && hoverIndex != g_dragTabIndex)
        {
            ReorderDocumentTab(g_dragTabIndex, hoverIndex);
            g_dragTabIndex = hoverIndex;
            return 0;
        }
        break;
    }
    case WM_MOUSELEAVE:
    {
        const int prevHoverIndex = g_hoverTabIndex;
        g_trackingTabsMouse = false;
        if (g_hoverTabIndex != -1 || g_hoverTabClose)
        {
            g_hoverTabIndex = -1;
            g_hoverTabClose = false;
            InvalidateTabItem(hwnd, prevHoverIndex);
        }
        break;
    }
    case WM_LBUTTONUP:
    {
        const bool wasDragging = g_draggingTab;
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
        if (wasDragging)
            return 0;
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
        if (reinterpret_cast<HWND>(lParam) != hwnd)
        {
            g_draggingTab = false;
            g_dragTabIndex = -1;
        }
        break;
    }
    return CallWindowProcW(g_origTabsProc, hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        BeginPaint(hwnd, &ps);
        PremiumOrchestrator::OnPaint(hwnd, ps.rcPaint);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        if (!hdc)
            return 0;
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const COLORREF bg = ThemeColorEditorBackground(IsDarkMode());
        HBRUSH hbr = CreateSolidBrush(bg);
        if (hbr)
        {
            FillRect(hdc, &rc, hbr);
            DeleteObject(hbr);
            return 1;
        }
        return 0;
    }
    case WM_MEASUREITEM:
    {
        auto *measure = reinterpret_cast<MEASUREITEMSTRUCT *>(lParam);
        if (HandlePopupMenuMeasureItem(measure))
            return TRUE;
        break;
    }
    case WM_DRAWITEM:
    {
        auto *draw = reinterpret_cast<DRAWITEMSTRUCT *>(lParam);
        if (HandlePopupMenuDrawItem(draw))
            return TRUE;
        break;
    }
    case WM_CREATE:
    {
        const auto &lang = GetLangStrings();
        g_hwndMain = hwnd;
        TabRefreshDpi();
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
                            lang.msgCannotLoadRichEdit.c_str(),
                            lang.msgError.c_str(),
                            MB_ICONERROR | MB_OK);
                return -1;
            }
        }
        g_editorClassName = richEditClass;
        const DWORD editorStyle = BuildEditorWindowStyle();
        g_hwndEditor = CreateWindowExW(0, richEditClass, nullptr,
                                       editorStyle,
                                       0, 0, 100, 100, hwnd, reinterpret_cast<HMENU>(IDC_EDITOR), GetModuleHandleW(nullptr), nullptr);
        g_origEditorProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndEditor, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(EditorSubclassProc)));
        
        UI::CustomScrollbar::RegisterClass(GetModuleHandle(nullptr));
        g_hwndScrollbar = UI::CustomScrollbar::Create(hwnd, g_hwndEditor);

        UI::SelectionAura::RegisterClass(GetModuleHandle(nullptr));
        g_hwndSelectionAura = UI::SelectionAura::Create(hwnd, g_hwndEditor);

        UI::CommandPalette::RegisterClass(GetModuleHandle(nullptr));
        g_hwndCommandPalette = UI::CommandPalette::Create(hwnd);
        
        // Populate Initial 'Studio' Commands
        UI::CommandPalette::AddCommand(L"File: Save", L"Save current document to disk", IDM_FILE_SAVE);
        UI::CommandPalette::AddCommand(L"File: Open", L"Open a document from disk", IDM_FILE_OPEN);
        UI::CommandPalette::AddCommand(L"Theme: Studio Charcoal", L"Switch to industrial dark theme", IDM_VIEW_THEME_CHARCOAL);
        UI::CommandPalette::AddCommand(L"Theme: Studio Paper", L"Switch to professional light theme", IDM_VIEW_THEME_PAPER);
        UI::CommandPalette::AddCommand(L"Editor: Zoom In", L"Increase font size", IDM_VIEW_ZOOMIN);
        UI::CommandPalette::AddCommand(L"Editor: Zoom Out", L"Decrease font size", IDM_VIEW_ZOOMOUT);

        ConfigureEditorControl(g_hwndEditor);
        if (kEnableCommandBar)
        {
            g_hwndCommandBar = CreateWindowExW(0, TOOLBARCLASSNAMEW, nullptr,
                                               WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TBSTYLE_FLAT | TBSTYLE_LIST | CCS_NOPARENTALIGN | CCS_NODIVIDER | CCS_NORESIZE,
                                               0, 0, 100, TabScalePx(DesignSystem::kChromeBandHeightPx), hwnd, reinterpret_cast<HMENU>(IDC_COMMANDBAR), GetModuleHandleW(nullptr), nullptr);
            if (g_hwndCommandBar)
            {
                SetWindowTheme(g_hwndCommandBar, L"Explorer", nullptr);
                InitializeCommandBar();
            }
        }
        g_hwndTabs = CreateWindowExW(0, WC_TABCONTROLW, L"",
                                     WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_TABSTOP | TCS_FOCUSNEVER | TCS_FIXEDWIDTH,
                                     0, 0, 100, TabScalePx(DesignSystem::kChromeBandHeightPx), hwnd, reinterpret_cast<HMENU>(IDC_TABS), GetModuleHandleW(nullptr), nullptr);
        if (g_hwndTabs)
        {
            TabRefreshDpi();
            TabRefreshVisualMetrics();
            SetWindowTheme(g_hwndTabs, L"", nullptr);
            g_tabsCustomDrawObserved = false;
            g_origTabsProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndTabs, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(TabsSubclassProc)));
        }
        g_hwndStatus = CreateWindowExW(0, STATUSCLASSNAMEW, nullptr,
                                       WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDC_STATUSBAR), GetModuleHandleW(nullptr), nullptr);
        g_origStatusProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndStatus, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(StatusSubclassProc)));
        RefreshChromeUiFont();
        SendMessageW(g_hwndEditor, EM_EXLIMITTEXT, 0, static_cast<LPARAM>(-1));
        SendMessageW(g_hwndEditor, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
        ApplyFont();
        SetupStatusBarParts();
        UpdateMenuStrings();
        UpdateRecentFilesMenu();
        UpdateLanguageMenu();
        RefreshCommandBarLabels();
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
        const bool dark = IsDarkMode();
        UAHMENU *pUDM = reinterpret_cast<UAHMENU *>(lParam);
        MENUBARINFO mbi = {};
        mbi.cbSize = sizeof(mbi);
        if (GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbi))
        {
            RECT rcWindow;
            GetWindowRect(hwnd, &rcWindow);
            RECT rcMenuBar = mbi.rcBar;
            OffsetRect(&rcMenuBar, -rcWindow.left, -rcWindow.top);
            const COLORREF menuBg = ThemeColorMenuBackground(dark);
            const COLORREF borderColor = ThemeColorChromeBorder(dark);
            HBRUSH hbrMenu = nullptr;
            if (dark && g_hbrMenuDark)
                hbrMenu = g_hbrMenuDark;
            else
                hbrMenu = CreateSolidBrush(menuBg);
            FillRect(pUDM->hdc, &rcMenuBar, hbrMenu ? hbrMenu : GetSysColorBrush(COLOR_MENU));
            if ((!dark || !g_hbrMenuDark) && hbrMenu)
                DeleteObject(hbrMenu);

            RECT separator = rcMenuBar;
            separator.top = std::max(separator.top, separator.bottom - 1);
            HBRUSH hbrBorder = CreateSolidBrush(borderColor);
            if (hbrBorder)
            {
                FillRect(pUDM->hdc, &separator, hbrBorder);
                DeleteObject(hbrBorder);
            }
            return TRUE;
        }
        break;
    }
    case WM_UAHDRAWMENUITEM:
    {
        const bool dark = IsDarkMode();
        UAHDRAWMENUITEM *pUDMI = reinterpret_cast<UAHDRAWMENUITEM *>(lParam);
        wchar_t szText[256] = {};
        MENUITEMINFOW mii = {};
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_STRING;
        mii.dwTypeData = szText;
        mii.cch = 255;
        GetMenuItemInfoW(pUDMI->um.hMenu, pUDMI->umi.iPosition, TRUE, &mii);
        COLORREF bgColor = ThemeColorMenuBackground(dark);
        COLORREF textColor = ThemeColorMenuText(dark);
        if ((pUDMI->dis.itemState & ODS_HOTLIGHT) || (pUDMI->dis.itemState & ODS_SELECTED))
            bgColor = ThemeColorMenuHoverBackground(dark);
        if (pUDMI->dis.itemState & ODS_DISABLED)
            textColor = ThemeColorMenuDisabledText(dark);
        HBRUSH hbr = CreateSolidBrush(bgColor);
        FillRect(pUDMI->um.hdc, &pUDMI->dis.rcItem, hbr);
        DeleteObject(hbr);
        HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(pUDMI->um.hdc, ChromeUiFontOrDefault()));
        SetBkMode(pUDMI->um.hdc, TRANSPARENT);
        SetTextColor(pUDMI->um.hdc, textColor);
        RECT rcText = MenuTextRectWithInsets(pUDMI->dis.rcItem);
        DrawTextW(pUDMI->um.hdc, szText, -1, &rcText, DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
        SelectObject(pUDMI->um.hdc, hOldFont);
        return TRUE;
        break;
    }
    case WM_NCPAINT:
    case WM_NCACTIVATE:
    {
        LRESULT result = DefWindowProcW(hwnd, msg, wParam, lParam);
        const bool dark = IsDarkMode();
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
            const COLORREF menuBg = ThemeColorMenuBackground(dark);
            const COLORREF menuText = ThemeColorMenuText(dark);
            const COLORREF borderColor = ThemeColorChromeBorder(dark);
            HBRUSH hbrMenu = nullptr;
            if (dark && g_hbrMenuDark)
                hbrMenu = g_hbrMenuDark;
            else
                hbrMenu = CreateSolidBrush(menuBg);
            FillRect(hdc, &rcMenuBar, hbrMenu ? hbrMenu : GetSysColorBrush(COLOR_MENU));
            if ((!dark || !g_hbrMenuDark) && hbrMenu)
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
            HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(hdc, ChromeUiFontOrDefault()));
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
                    RECT rcText = MenuTextRectWithInsets(rcItem);
                    DrawTextW(hdc, szText, -1, &rcText, DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
                }
            }
            SelectObject(hdc, hOldFont);
        }
        ReleaseDC(hwnd, hdc);
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
        PremiumOrchestrator::OnSize(static_cast<UINT>(LOWORD(lParam)), static_cast<UINT>(HIWORD(lParam)));
        TabRefreshDpi();
        ResizeControls();
        UpdateStatus();
        return 0;
    case WM_DPICHANGED:
    {
        const RECT *suggested = reinterpret_cast<const RECT *>(lParam);
        if (suggested)
        {
            SetWindowPos(hwnd, nullptr, suggested->left, suggested->top,
                         suggested->right - suggested->left,
                         suggested->bottom - suggested->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
        TabRefreshDpi();
        TabRefreshVisualMetrics();
        g_tabsCustomDrawObserved = false;
        RefreshChromeUiFont();
        ApplyEditorScrollbarChrome();
        RefreshCommandBarMetrics();
        PremiumOrchestrator::OnDpiChanged(hwnd);
        ResizeControls();
        UpdateStatus();
        return 0;
    }
    case WM_TIMER:
        if (PremiumOrchestrator::OnTimer(hwnd, wParam))
            return 0;
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
        const std::wstring preFilePath = g_state.filePath;
        const bool preModified = g_state.modified;
        const Encoding preEncoding = g_state.encoding;
        const LineEnding preLineEnding = g_state.lineEnding;
        const bool preLargeFileMode = g_state.largeFileMode;
        const size_t preLargeFileBytes = g_state.largeFileBytes;
        const size_t preDocumentCount = g_documents.size();
        const int preActiveDocument = g_activeDocument;
        if (cmd == IDC_EDITOR && HIWORD(wParam) == EN_CHANGE)
        {
            if (g_switchingDocument)
                return 0;
            if (!g_state.modified)
            {
                g_state.modified = true;
                UpdateTitle();
            }
#ifdef EM_SHOWSCROLLBAR
            SendMessageW(g_hwndEditor, EM_SHOWSCROLLBAR, SB_VERT, FALSE);
            SendMessageW(g_hwndEditor, EM_SHOWSCROLLBAR, SB_HORZ, FALSE);
#endif
            ShowScrollBar(g_hwndEditor, SB_VERT, FALSE);
            ShowScrollBar(g_hwndEditor, SB_HORZ, FALSE);
            if (g_hwndScrollbar)
                InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
            SyncDocumentFromState(g_activeDocument, false);
            RefreshCommandBarStates();
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
        default:
            if (RouteStandardCommand(hwnd, cmd))
                break;
            switch (cmd)
            {
            case IDM_VIEW_COMMAND_PALETTE:
                UI::CommandPalette::Show(g_hwndCommandPalette);
                break;
            case IDM_VIEW_THEME_CHARCOAL:
                g_state.theme = Theme::Dark;
                ApplyTheme();
                break;
            case IDM_VIEW_THEME_PAPER:
                g_state.theme = Theme::Light;
                ApplyTheme();
                break;
            case IDM_VIEW_USETABS:
                if (g_state.useTabs && g_documents.size() > 1)
                {
                    const auto &lang = GetLangStrings();
                    MessageBoxW(hwnd,
                                lang.msgDisableTabsCloseOthers.c_str(),
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
                RefreshCommandBarLabels();
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
                RefreshCommandBarLabels();
                RefreshAllDocumentTabLabels();
                UpdateTitle();
                UpdateStatus();
                UpdateRuntimeMenuStates();
                break;
            default:
                break;
            }
            break;
        }
        SyncDocumentFromState(g_activeDocument, false);
        const bool sessionStateChanged =
            (g_documents.size() != preDocumentCount) ||
            (g_activeDocument != preActiveDocument) ||
            (g_state.filePath != preFilePath) ||
            (g_state.modified != preModified) ||
            (g_state.encoding != preEncoding) ||
            (g_state.lineEnding != preLineEnding) ||
            (g_state.largeFileMode != preLargeFileMode) ||
            (g_state.largeFileBytes != preLargeFileBytes);
        if (sessionStateChanged && !g_state.closing)
            MarkSessionDirty();
        return 0;
    }
    case WM_NOTIFY:
    {
        NMHDR *pnmh = reinterpret_cast<NMHDR *>(lParam);
        if (pnmh->hwndFrom == g_hwndCommandBar && pnmh->code == NM_CUSTOMDRAW)
            return HandleCommandBarCustomDraw(reinterpret_cast<LPNMTBCUSTOMDRAW>(lParam));
        if (pnmh->hwndFrom == g_hwndTabs && !g_state.useTabs)
            return 0;
        if (pnmh->hwndFrom == g_hwndTabs && pnmh->code == NM_CUSTOMDRAW)
        {
            // Tab strip rendering is fully handled in subclass WM_PAINT for both themes.
            return CDRF_DODEFAULT;
        }
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
            AppendMenuW(hPopup, MF_STRING, IDM_FILE_CLOSETAB, lang.menuCloseTab.c_str());
            AppendMenuW(hPopup, MF_STRING, IDM_FILE_REOPENCLOSEDTAB, lang.menuReopenClosedTab.c_str());
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
        if (pnmh->hwndFrom == g_hwndEditor && pnmh->code == EN_SELCHANGE)
        {
            if (g_hwndSelectionAura) UI::SelectionAura::Update(g_hwndSelectionAura);
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
        PremiumOrchestrator::Shutdown(hwnd);
        TabSpinDetach();
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
        ShutdownBackgroundGraphics();
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
        if (g_hChromeUiFont)
        {
            DeleteObject(g_hChromeUiFont);
            g_hChromeUiFont = nullptr;
        }
        if (g_hRichEditModule)
        {
            FreeLibrary(g_hRichEditModule);
            g_hRichEditModule = nullptr;
        }
        UnloadBundledFonts();
        TabDestroyFonts();
        ReleaseTabBackbuffer();
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
        if (g_hwndEditor && ScrollEditorFromMouseWheel(g_hwndEditor, wParam))
        {
            if (g_hwndScrollbar)
                InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
            return 0;
        }
        break;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow)
{
    InitializeCrashDiagnostics();
    CrashDiagnosticsLog(L"wWinMain started");
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

    LoadBundledFonts();
    LoadFontSettings();
    ApplyRuntimeFeatureOverrides();

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

    INITCOMMONCONTROLSEX icc = {sizeof(icc), ICC_BAR_CLASSES};
    InitCommonControlsEx(&icc);

    const auto &lang = GetLangStrings();
    std::wstring initialTitle = lang.untitled + L" - " + lang.appName;
    g_hwndMain = CreateWindowExW(0, L"NotepadClass", initialTitle.c_str(),
                                 WS_OVERLAPPEDWINDOW | WS_MAXIMIZEBOX | WS_CLIPCHILDREN, g_state.windowX, g_state.windowY, g_state.windowWidth, g_state.windowHeight,
                                 nullptr, nullptr, hInstance, nullptr);
    CrashDiagnosticsLog(L"Main window created");
    g_hAccel = LoadAcceleratorsW(hInstance, MAKEINTRESOURCEW(IDR_ACCEL));
    int showCmd = benchmarkOnly ? SW_HIDE : nCmdShow;
    if (!benchmarkOnly && g_state.windowMaximized && (nCmdShow == SW_SHOW || nCmdShow == SW_SHOWNORMAL || nCmdShow == SW_SHOWDEFAULT))
        showCmd = SW_SHOWMAXIMIZED;
    ShowWindow(g_hwndMain, showCmd);
    UpdateWindow(g_hwndMain);
    PremiumOrchestrator::Initialize(g_hwndMain);

    if (benchmarkOnly)
    {
        bool allPassed = false;
        bool allExecuted = false;
        const bool ok = RunPerformanceBenchmark(false, nullptr, &allPassed, &allExecuted);
        CrashDiagnosticsLog(ok ? L"Benchmark mode exit success" : L"Benchmark mode exit failure");
        ShutdownBackgroundGraphics();
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
        // When Command Palette is active, route ALL keyboard input
        // directly to it — bypass accelerators and dialog hooks.
        if (g_hwndCommandPalette && UI::CommandPalette::IsVisible(g_hwndCommandPalette))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
            continue;
        }
        if (g_hwndFindDlg && IsDialogMessageW(g_hwndFindDlg, &msg))
            continue;
        if (!TranslateAcceleratorW(g_hwndMain, g_hAccel, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    CrashDiagnosticsLog(L"Application message loop exited");
    ShutdownBackgroundGraphics();
    return static_cast<int>(msg.wParam);
}
