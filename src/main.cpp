/*
  Saka Studio & Engineering

  Main entry point and window procedure for Legacy Notepad text editor application.
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
};

static std::vector<DocumentTabState> g_documents;
static int g_activeDocument = -1;
static bool g_switchingDocument = false;

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

static void SyncDocumentFromState(int index, bool includeText)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    DocumentTabState &doc = g_documents[index];
    if (includeText)
        doc.text = GetEditorText();
    doc.filePath = g_state.filePath;
    doc.modified = g_state.modified;
    doc.encoding = g_state.encoding;
    doc.lineEnding = g_state.lineEnding;
    SetDocumentTabLabel(index);
}

static void LoadStateFromDocument(int index)
{
    if (index < 0 || index >= static_cast<int>(g_documents.size()))
        return;

    g_switchingDocument = true;
    const DocumentTabState &doc = g_documents[index];
    g_state.filePath = doc.filePath;
    g_state.modified = doc.modified;
    g_state.encoding = doc.encoding;
    g_state.lineEnding = doc.lineEnding;
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

    g_documents.push_back(doc);
    g_activeDocument = 0;

    TCITEMW item{};
    item.mask = TCIF_TEXT;
    std::wstring label = DocumentTabLabel(doc);
    item.pszText = label.data();
    TabCtrl_InsertItem(g_hwndTabs, 0, &item);
    TabCtrl_SetCurSel(g_hwndTabs, 0);
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

static bool OpenPathInTabs(const std::wstring &path)
{
    if (path.empty())
        return false;

    const int existingIndex = FindDocumentByPath(path);
    if (existingIndex >= 0)
    {
        SwitchToDocument(existingIndex);
        return true;
    }

    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    if (!IsCurrentDocumentEmptyAndUntitled())
        CreateNewDocumentTab();

    LoadFile(path);
    SyncDocumentFromState(g_activeDocument, true);
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

static void CloseCurrentDocumentTab()
{
    if (g_activeDocument < 0 || g_activeDocument >= static_cast<int>(g_documents.size()))
        return;

    if (g_documents.size() <= 1)
    {
        if (!ConfirmDiscard())
            return;
        FileNew();
        SyncDocumentFromState(g_activeDocument, true);
        return;
    }

    if (!ConfirmDiscard())
        return;

    const int closingIndex = g_activeDocument;
    g_documents.erase(g_documents.begin() + closingIndex);
    if (g_hwndTabs)
        TabCtrl_DeleteItem(g_hwndTabs, closingIndex);

    int nextIndex = closingIndex;
    if (nextIndex >= static_cast<int>(g_documents.size()))
        nextIndex = static_cast<int>(g_documents.size()) - 1;

    g_activeDocument = nextIndex;
    if (g_hwndTabs)
        TabCtrl_SetCurSel(g_hwndTabs, nextIndex);
    LoadStateFromDocument(nextIndex);
    RefreshAllDocumentTabLabels();
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

static bool ConfirmDiscardAllDocuments()
{
    if (g_documents.empty())
        return ConfirmDiscard();

    if (g_activeDocument >= 0)
        SyncDocumentFromState(g_activeDocument, true);

    const int originalIndex = g_activeDocument;
    for (int i = 0; i < static_cast<int>(g_documents.size()); ++i)
    {
        if (g_activeDocument != i)
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

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        g_hwndMain = hwnd;
        DragAcceptFiles(hwnd, TRUE);
        const wchar_t *richEditClass = nullptr;
        HMODULE hRichEdit = nullptr;
        hRichEdit = LoadLibraryW(L"Msftedit.dll");
        if (hRichEdit)
        {
            richEditClass = MSFTEDIT_CLASS;
        }
        else
        {
            hRichEdit = LoadLibraryW(L"Riched20.dll");
            if (hRichEdit)
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
        g_hwndEditor = CreateWindowExW(0, richEditClass, nullptr,
                                       WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN | ES_NOHIDESEL,
                                       0, 0, 100, 100, hwnd, reinterpret_cast<HMENU>(IDC_EDITOR), GetModuleHandleW(nullptr), nullptr);
        g_origEditorProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndEditor, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(EditorSubclassProc)));
        ConfigureEditorControl(g_hwndEditor);
        g_hwndTabs = CreateWindowExW(0, WC_TABCONTROLW, L"",
                                     WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_TABSTOP,
                                     0, 0, 100, 30, hwnd, reinterpret_cast<HMENU>(IDC_TABS), GetModuleHandleW(nullptr), nullptr);
        if (g_hwndTabs)
        {
            HFONT hFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            SendMessageW(g_hwndTabs, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
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
        if (g_state.wordWrap)
            ApplyWordWrap();
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
                FillRect(pUDM->hdc, &rcMenuBar, g_hbrMenuDark ? g_hbrMenuDark : reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
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
            COLORREF bgColor = RGB(45, 45, 45);
            COLORREF textColor = RGB(255, 255, 255);
            if ((pUDMI->dis.itemState & ODS_HOTLIGHT) || (pUDMI->dis.itemState & ODS_SELECTED))
                bgColor = RGB(65, 65, 65);
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
                rcMenuBar.bottom += 2;
                FillRect(hdc, &rcMenuBar, g_hbrMenuDark ? g_hbrMenuDark : reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
                HMENU hMenu = GetMenu(hwnd);
                int itemCount = GetMenuItemCount(hMenu);
                NONCLIENTMETRICSW ncm = {};
                ncm.cbSize = sizeof(ncm);
                SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
                HFONT hFont = CreateFontIndirectW(&ncm.lfMenuFont);
                HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(hdc, hFont));
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, RGB(255, 255, 255));
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
        for (UINT i = 0; i < fileCount; ++i)
        {
            wchar_t path[MAX_PATH] = {};
            if (DragQueryFileW(hDrop, i, path, MAX_PATH))
                OpenPathInTabs(path);
        }
        DragFinish(hDrop);
        return 0;
    }
    case WM_SIZE:
        ResizeControls();
        UpdateStatus();
        return 0;
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
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return reinterpret_cast<LRESULT>(g_hbrStatusDark ? g_hbrStatusDark : GetStockObject(BLACK_BRUSH));
        }
        break;
    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT pDIS = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);
        if (pDIS->hwndItem == g_hwndStatus && IsDarkMode())
        {
            HBRUSH hbr = g_hbrStatusDark ? g_hbrStatusDark : CreateSolidBrush(RGB(45, 45, 45));
            FillRect(pDIS->hDC, &pDIS->rcItem, hbr);
            if (!g_hbrStatusDark)
                DeleteObject(hbr);
            SetBkMode(pDIS->hDC, TRANSPARENT);
            SetTextColor(pDIS->hDC, RGB(255, 255, 255));
            int part = static_cast<int>(pDIS->itemID);
            if (part >= 0 && part < 4)
            {
                RECT rc = pDIS->rcItem;
                rc.left += 4;
                DrawTextW(pDIS->hDC, g_statusTexts[part].c_str(), -1, &rc, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);
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

        const UINT cmd = TrackPopupMenu(hPopup, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, nullptr);
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
            CreateNewDocumentTab();
            break;
        case IDM_FILE_OPEN:
            OpenFileInNewDocumentTabDialog();
            break;
        case IDM_FILE_CLOSETAB:
            CloseCurrentDocumentTab();
            break;
        case IDM_FILE_NEXTTAB:
            SwitchToNextDocumentTab(false);
            break;
        case IDM_FILE_PREVTAB:
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
            if (ConfirmDiscardAllDocuments())
                DestroyWindow(hwnd);
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
        case IDM_HELP_ABOUT:
            HelpAbout();
            break;
        }
        SyncDocumentFromState(g_activeDocument, false);
        return 0;
    }
    case WM_NOTIFY:
    {
        NMHDR *pnmh = reinterpret_cast<NMHDR *>(lParam);
        if (pnmh->hwndFrom == g_hwndTabs && pnmh->code == TCN_SELCHANGE)
        {
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
            if (g_documents.size() <= 1)
                EnableMenuItem(hPopup, IDM_FILE_CLOSETAB, MF_BYCOMMAND | MF_GRAYED);

            UINT cmd = TrackPopupMenu(hPopup, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD,
                                      pt.x, pt.y, 0, hwnd, nullptr);
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
                if (lpnmcd->dwDrawStage == CDDS_PREPAINT)
                    return CDRF_NOTIFYITEMDRAW;
                if (lpnmcd->dwDrawStage == CDDS_ITEMPREPAINT)
                {
                    HBRUSH hbr = g_hbrStatusDark ? g_hbrStatusDark : CreateSolidBrush(RGB(45, 45, 45));
                    FillRect(lpnmcd->hdc, &lpnmcd->rc, hbr);
                    if (!g_hbrStatusDark && hbr)
                        DeleteObject(hbr);
                    SetBkMode(lpnmcd->hdc, TRANSPARENT);
                    SetBkColor(lpnmcd->hdc, RGB(45, 45, 45));
                    SetTextColor(lpnmcd->hdc, RGB(255, 255, 255));
                    wchar_t buf[256] = {};
                    int part = static_cast<int>(lpnmcd->dwItemSpec);
                    SendMessageW(g_hwndStatus, SB_GETTEXTW, part, reinterpret_cast<LPARAM>(buf));
                    RECT rc = lpnmcd->rc;
                    rc.left += 6;
                    DrawTextW(lpnmcd->hdc, buf, -1, &rc, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);
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
        g_state.closing = true;
        if (ConfirmDiscardAllDocuments())
            DestroyWindow(hwnd);
        else
            g_state.closing = false;
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

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR lpCmdLine, int nCmdShow)
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
    int showCmd = nCmdShow;
    if (g_state.windowMaximized && (nCmdShow == SW_SHOW || nCmdShow == SW_SHOWNORMAL || nCmdShow == SW_SHOWDEFAULT))
        showCmd = SW_SHOWMAXIMIZED;
    ShowWindow(g_hwndMain, showCmd);
    UpdateWindow(g_hwndMain);

    if (lpCmdLine && lpCmdLine[0])
    {
        std::wstring path = lpCmdLine;
        if (path.front() == L'"' && path.back() == L'"')
            path = path.substr(1, path.size() - 2);
        OpenPathInTabs(path);
    }

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
