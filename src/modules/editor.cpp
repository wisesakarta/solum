/*
  Technical Standard

  Editor control functions for text manipulation, font rendering, and zoom control.
  Handles RichEdit control subclassing, word wrap, and cursor position tracking.
*/

#include "editor.h"
#include "core/list_continuation.h"
#include "core/types.h"
#include "core/globals.h"
#include "theme.h"
#include "ui.h"
#include "background.h"
#include "design_system.h"
#include "custom_scrollbar.h"
#include "selection_aura.h"
#include "resource.h"
#include <richedit.h>
#include <commctrl.h>
#include <algorithm>
#include <cwctype>

#ifndef EM_BEGINUNDOACTION
#define EM_BEGINUNDOACTION (WM_USER + 84)
#endif

#ifndef EM_ENDUNDOACTION
#define EM_ENDUNDOACTION (WM_USER + 85)
#endif

struct StreamCookie
{
    const std::wstring *text;
    size_t pos;
};

static DWORD CALLBACK StreamInCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamCookie *pCookie = reinterpret_cast<StreamCookie *>(dwCookie);
    size_t remaining = (pCookie->text->length() * sizeof(wchar_t)) - pCookie->pos;
    if (remaining <= 0)
    {
        *pcb = 0;
        return 0;
    }
    size_t toCopy = (static_cast<size_t>(cb) < remaining) ? static_cast<size_t>(cb) : remaining;
    memcpy(pbBuff, reinterpret_cast<const BYTE *>(pCookie->text->c_str()) + pCookie->pos, toCopy);
    pCookie->pos += toCopy;
    *pcb = static_cast<LONG>(toCopy);
    return 0;
}

struct StreamOutCookie
{
    std::wstring *text;
};

static DWORD CALLBACK StreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamOutCookie *pCookie = reinterpret_cast<StreamOutCookie *>(dwCookie);
    pCookie->text->append(reinterpret_cast<const wchar_t *>(pbBuff), cb / sizeof(wchar_t));
    *pcb = cb;
    return 0;
}

struct StreamInAnsiCookie
{
    const std::string *text;
    size_t pos;
};

static DWORD CALLBACK StreamInAnsiCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamInAnsiCookie *pCookie = reinterpret_cast<StreamInAnsiCookie *>(dwCookie);
    size_t remaining = pCookie->text->size() - pCookie->pos;
    if (remaining == 0)
    {
        *pcb = 0;
        return 0;
    }
    size_t toCopy = (static_cast<size_t>(cb) < remaining) ? static_cast<size_t>(cb) : remaining;
    memcpy(pbBuff, pCookie->text->data() + pCookie->pos, toCopy);
    pCookie->pos += toCopy;
    *pcb = static_cast<LONG>(toCopy);
    return 0;
}

struct StreamOutAnsiCookie
{
    std::string *text;
};

static DWORD CALLBACK StreamOutAnsiCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamOutAnsiCookie *pCookie = reinterpret_cast<StreamOutAnsiCookie *>(dwCookie);
    pCookie->text->append(reinterpret_cast<const char *>(pbBuff), static_cast<size_t>(cb));
    *pcb = cb;
    return 0;
}

static bool IsEditorWrapEnabled()
{
    return g_state.wordWrap && !g_state.largeFileMode;
}

DWORD BuildEditorWindowStyle()
{
    DWORD style = WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN | ES_NOHIDESEL;
    if (!IsEditorWrapEnabled())
        style |= WS_HSCROLL | ES_AUTOHSCROLL;
    return style;
}

static int ScaleEditorPx(int px)
{
    const HWND ref = g_hwndEditor ? g_hwndEditor : g_hwndMain;
    return DesignSystem::ScalePx(px, ref);
}

static void ApplyFlatScrollbarStyle(HWND hwnd)
{
    if (!hwnd)
        return;

    InitializeFlatSB(hwnd);

    const bool dark = IsDarkMode();
    const int barThickness = ScaleEditorPx(14);
    const int thumbExtent = ScaleEditorPx(26);
    const COLORREF trackColor = ThemeColorEditorBackground(dark);

    FlatSB_SetScrollProp(hwnd, WSB_PROP_VSTYLE, FSB_FLAT_MODE, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_HSTYLE, FSB_FLAT_MODE, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_CXVSCROLL, barThickness, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_CYHSCROLL, barThickness, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_CYVTHUMB, thumbExtent, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_CXHTHUMB, thumbExtent, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_VBKGCOLOR, trackColor, FALSE);
    FlatSB_SetScrollProp(hwnd, WSB_PROP_HBKGCOLOR, trackColor, FALSE);
    InvalidateRect(hwnd, nullptr, FALSE);
}

static void ConfigureEditorInputExperience(HWND hwnd)
{
    if (!hwnd)
        return;

    DWORD disableLangOptions = 0;
#ifdef IMF_SPELLCHECKING
    disableLangOptions |= IMF_SPELLCHECKING;
#endif
#ifdef IMF_TKBPREDICTION
    disableLangOptions |= IMF_TKBPREDICTION;
#endif

    if (disableLangOptions != 0)
    {
        const LRESULT currentLangOptions = SendMessageW(hwnd, EM_GETLANGOPTIONS, 0, 0);
        SendMessageW(hwnd, EM_SETLANGOPTIONS, 0, static_cast<LPARAM>(currentLangOptions & ~static_cast<LRESULT>(disableLangOptions)));
    }

    DWORD disableEditStyles = 0;
#ifdef SES_CTFALLOWPROOFING
    disableEditStyles |= SES_CTFALLOWPROOFING;
#endif
#ifdef SES_CTFALLOWSMARTTAG
    disableEditStyles |= SES_CTFALLOWSMARTTAG;
#endif
#ifdef SES_CTFALLOWEMBED
    disableEditStyles |= SES_CTFALLOWEMBED;
#endif

    if (disableEditStyles != 0)
        SendMessageW(hwnd, EM_SETEDITSTYLE, disableEditStyles, 0);

#ifdef EM_SETAUTOCORRECTPROC
    SendMessageW(hwnd, EM_SETAUTOCORRECTPROC, 0, 0);
#endif
}

static void EnsureNativeScrollbarsHidden(HWND hwnd)
{
    if (!hwnd)
        return;
#ifdef EM_SHOWSCROLLBAR
    SendMessageW(hwnd, EM_SHOWSCROLLBAR, SB_VERT, FALSE);
    SendMessageW(hwnd, EM_SHOWSCROLLBAR, SB_HORZ, FALSE);
#endif
    ShowScrollBar(hwnd, SB_VERT, FALSE);
    ShowScrollBar(hwnd, SB_HORZ, FALSE);
}

namespace
{
bool g_suppressNextReturnChar = false;
int g_verticalWheelDeltaRemainder = 0;

struct ListContinuation
{
    bool enabled = false;
    bool exitListMode = false;
    LONG lineStart = 0;
    LONG lineEnd = 0;
    std::wstring continuationPrefix;
};

bool BuildListContinuationForEnter(HWND hwnd, ListContinuation &continuation)
{
    DWORD selStart = 0;
    DWORD selEnd = 0;
    SendMessageW(hwnd, EM_GETSEL, reinterpret_cast<WPARAM>(&selStart), reinterpret_cast<LPARAM>(&selEnd));
    if (selStart != selEnd)
        return false;

    const LONG lineIndex = static_cast<LONG>(SendMessageW(hwnd, EM_EXLINEFROMCHAR, 0, static_cast<LPARAM>(selStart)));
    if (lineIndex < 0)
        return false;

    const LONG lineStart = static_cast<LONG>(SendMessageW(hwnd, EM_LINEINDEX, static_cast<WPARAM>(lineIndex), 0));
    if (lineStart < 0)
        return false;

    const LONG lineLength = static_cast<LONG>(SendMessageW(hwnd, EM_LINELENGTH, static_cast<WPARAM>(lineStart), 0));
    if (lineLength < 0)
        return false;

    const LONG logicalLineEnd = lineStart + lineLength;
    continuation.lineStart = lineStart;
    continuation.lineEnd = logicalLineEnd;
    const LONG caretInLine = static_cast<LONG>(selStart) - lineStart;

    std::wstring lineText(static_cast<size_t>(lineLength) + 1, L'\0');
    if (lineLength > 0)
    {
        TEXTRANGEW textRange{};
        textRange.chrg.cpMin = lineStart;
        textRange.chrg.cpMax = logicalLineEnd;
        textRange.lpstrText = lineText.data();
        SendMessageW(hwnd, EM_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&textRange));
    }
    lineText.resize(static_cast<size_t>(lineLength));

    const ListContinuationPlan plan = BuildListContinuationPlan(lineText, static_cast<size_t>(caretInLine));
    if (!plan.matched)
        return false;

    continuation.enabled = true;
    continuation.exitListMode = plan.exitListMode;
    continuation.continuationPrefix = plan.continuationPrefix;
    return true;
}

bool HandleAutoListEnterKeyDown(HWND hwnd, WPARAM wParam, LPARAM)
{
    if (wParam != VK_RETURN)
        return false;
    if ((GetKeyState(VK_CONTROL) & 0x8000) || (GetKeyState(VK_MENU) & 0x8000) || (GetKeyState(VK_SHIFT) & 0x8000))
        return false;

    ListContinuation continuation{};
    if (!BuildListContinuationForEnter(hwnd, continuation) || !continuation.enabled)
        return false;

    if (continuation.exitListMode)
    {
        SendMessageW(hwnd, EM_BEGINUNDOACTION, 0, 0);
        SendMessageW(hwnd, EM_SETSEL, static_cast<WPARAM>(continuation.lineStart), static_cast<LPARAM>(continuation.lineEnd));
        SendMessageW(hwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
        SendMessageW(hwnd, EM_ENDUNDOACTION, 0, 0);
        g_suppressNextReturnChar = true;
        return true;
    }

    std::wstring insertText = L"\r\n";
    insertText += continuation.continuationPrefix;

    SendMessageW(hwnd, EM_BEGINUNDOACTION, 0, 0);
    SendMessageW(hwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(insertText.c_str()));
    SendMessageW(hwnd, EM_ENDUNDOACTION, 0, 0);
    g_suppressNextReturnChar = true;
    return true;
}

bool IsKeyPressed(int virtualKey)
{
    return (GetKeyState(virtualKey) & 0x8000) != 0;
}

bool HandleInlineFormattingShortcut(HWND hwnd, WPARAM wParam)
{
    if (!g_hwndMain || !IsKeyPressed(VK_CONTROL) || IsKeyPressed(VK_MENU))
        return false;

    const bool shiftPressed = IsKeyPressed(VK_SHIFT);
    WORD commandId = 0;
    switch (wParam)
    {
    case 'B':
        if (!shiftPressed)
            commandId = IDM_FORMAT_BOLD;
        break;
    case 'I':
        if (!shiftPressed)
            commandId = IDM_FORMAT_ITALIC;
        break;
    case 'X':
        if (shiftPressed)
            commandId = IDM_FORMAT_STRIKETHROUGH;
        break;
    default:
        break;
    }

    if (commandId == 0)
        return false;

    SendMessageW(g_hwndMain, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
    SetFocus(hwnd);
    return true;
}

bool HandleTabManagementShortcut(HWND hwnd, WPARAM wParam)
{
    if (!g_hwndMain || !g_state.useTabs || !IsKeyPressed(VK_CONTROL) || IsKeyPressed(VK_MENU))
        return false;

    const bool shiftPressed = IsKeyPressed(VK_SHIFT);
    WORD commandId = 0;
    switch (wParam)
    {
    case VK_TAB:
        commandId = shiftPressed ? IDM_FILE_PREVTAB : IDM_FILE_NEXTTAB;
        break;
    case VK_NEXT:
        commandId = IDM_FILE_NEXTTAB; // Ctrl+PgDn
        break;
    case VK_PRIOR:
        commandId = IDM_FILE_PREVTAB; // Ctrl+PgUp
        break;
    case 'W':
        if (!shiftPressed)
            commandId = IDM_FILE_CLOSETAB;
        break;
    case 'T':
        if (shiftPressed)
            commandId = IDM_FILE_REOPENCLOSEDTAB;
        break;
    default:
        break;
    }

    if (commandId == 0)
        return false;

    SendMessageW(g_hwndMain, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
    SetFocus(hwnd);
    return true;
}
}

std::wstring GetEditorText()
{
    std::wstring text;
    StreamOutCookie cookie = {&text};
    EDITSTREAM es = {reinterpret_cast<DWORD_PTR>(&cookie), 0, StreamOutCallback};
    SendMessageW(g_hwndEditor, EM_STREAMOUT, SF_TEXT | SF_UNICODE, reinterpret_cast<LPARAM>(&es));
    return text;
}

void SetEditorText(const std::wstring &text)
{
    StreamCookie cookie = {&text, 0};
    EDITSTREAM es = {reinterpret_cast<DWORD_PTR>(&cookie), 0, StreamInCallback};
    SendMessageW(g_hwndEditor, EM_STREAMIN, SF_TEXT | SF_UNICODE, reinterpret_cast<LPARAM>(&es));
}

std::string GetEditorRichText()
{
    std::string rtf;
    StreamOutAnsiCookie cookie = {&rtf};
    EDITSTREAM es = {reinterpret_cast<DWORD_PTR>(&cookie), 0, StreamOutAnsiCallback};
    SendMessageW(g_hwndEditor, EM_STREAMOUT, SF_RTF, reinterpret_cast<LPARAM>(&es));
    return rtf;
}

void SetEditorRichText(const std::string &rtf)
{
    if (rtf.empty())
    {
        SetEditorText(L"");
        return;
    }

    StreamInAnsiCookie cookie = {&rtf, 0};
    EDITSTREAM es = {reinterpret_cast<DWORD_PTR>(&cookie), 0, StreamInAnsiCallback};
    SendMessageW(g_hwndEditor, EM_STREAMIN, SF_RTF, reinterpret_cast<LPARAM>(&es));
}

std::pair<int, int> GetCursorPos()
{
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    int line = static_cast<int>(SendMessageW(g_hwndEditor, EM_EXLINEFROMCHAR, 0, start));
    int lineIndex = static_cast<int>(SendMessageW(g_hwndEditor, EM_LINEINDEX, static_cast<WPARAM>(line), 0));
    int col = static_cast<int>(start) - lineIndex;
    return {line + 1, col + 1};
}

void SetEditorLineSpacing(HWND hwnd, float multiplier)
{
    if (!hwnd)
        return;
    PARAFORMAT2 pf = {};
    pf.cbSize = sizeof(pf);
    pf.dwMask = PFM_LINESPACING;
    pf.bLineSpacingRule = 5; // Spacing is dyLineSpacing / 20
    pf.dyLineSpacing = static_cast<LONG>(20 * multiplier);
    SendMessageW(hwnd, EM_SETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&pf));
}

void ConfigureEditorControl(HWND hwnd)
{
    if (!hwnd)
        return;

    // Use rich text mode for inline formatting commands (bold/italic/strikethrough),
    // while keeping large-file mode plain-text-first for responsiveness.
    SendMessageW(hwnd, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L""));
    const DWORD mode = (g_state.largeFileMode ? TM_PLAINTEXT : TM_RICHTEXT) | TM_MULTILEVELUNDO | TM_MULTICODEPAGE;
    if (SendMessageW(hwnd, EM_SETTEXTMODE, mode, 0) != 0)
        SendMessageW(hwnd, EM_SETTEXTMODE, TM_PLAINTEXT | TM_MULTILEVELUNDO | TM_MULTICODEPAGE, 0);
    SendMessageW(hwnd, EM_AUTOURLDETECT, FALSE, 0);
    ConfigureEditorInputExperience(hwnd);
    ApplyFlatScrollbarStyle(hwnd);
    EnsureNativeScrollbarsHidden(hwnd);

    SetEditorLineSpacing(hwnd, 1.2f); // Renaissance Rhythm (1.2x)
}

void ApplyEditorViewportPadding()
{
    if (!g_hwndEditor)
        return;

    RECT rc{};
    GetClientRect(g_hwndEditor, &rc);
    if (rc.right <= rc.left || rc.bottom <= rc.top)
        return;

    int padLeft = ScaleEditorPx(8);
    int padRight = ScaleEditorPx(8);
    int padTop = ScaleEditorPx(12);
    int padBottom = ScaleEditorPx(6);

    if ((rc.right - rc.left) <= (padLeft + padRight + 4))
    {
        padLeft = 2;
        padRight = 2;
    }
    if ((rc.bottom - rc.top) <= (padTop + padBottom + 4))
    {
        padTop = 2;
        padBottom = 2;
    }

    RECT formatRect{};
    formatRect.left = rc.left + padLeft;
    formatRect.top = rc.top + padTop;
    formatRect.right = rc.right - padRight;
    formatRect.bottom = rc.bottom - padBottom;

    if (formatRect.right <= formatRect.left)
        formatRect.right = formatRect.left + 1;
    if (formatRect.bottom <= formatRect.top)
        formatRect.bottom = formatRect.top + 1;

    SendMessageW(g_hwndEditor, EM_SETRECTNP, 0, reinterpret_cast<LPARAM>(&formatRect));
}

void ApplyEditorScrollbarChrome()
{
    ApplyFlatScrollbarStyle(g_hwndEditor);
}

void ApplyFont()
{
    if (!g_hwndEditor)
        return;
    if (g_state.hFont)
    {
        DeleteObject(g_state.hFont);
        g_state.hFont = nullptr;
    }
    int size = g_state.fontSize * g_state.zoomLevel / 100;
    size = (size < 8) ? 8 : (size > 500) ? 500
                                         : size;
    HDC hdc = GetDC(g_hwndMain);
    const int dpiY = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
    if (hdc)
        ReleaseDC(g_hwndMain, hdc);
    int height = -MulDiv(size, (dpiY > 0) ? dpiY : 96, 72);
    auto createWithFace = [&](const wchar_t *face) -> HFONT
    {
        return CreateFontW(height, 0, 0, 0, g_state.fontWeight, g_state.fontItalic, g_state.fontUnderline, FALSE,
                           DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                           DEFAULT_PITCH | FF_DONTCARE, face);
    };

    g_state.hFont = createWithFace(g_state.fontName.c_str());
    if (!g_state.hFont && lstrcmpiW(g_state.fontName.c_str(), DesignSystem::kUiFontPrimary) != 0)
        g_state.hFont = createWithFace(DesignSystem::kUiFontPrimary);
    if (!g_state.hFont && lstrcmpiW(g_state.fontName.c_str(), DesignSystem::kUiFontFallback) != 0)
        g_state.hFont = createWithFace(DesignSystem::kUiFontFallback);
    if (!g_state.hFont)
        g_state.hFont = createWithFace(L"Consolas");
    if (!g_state.hFont)
        return;

    LOGFONTW appliedLf{};
    if (GetObjectW(g_state.hFont, sizeof(appliedLf), &appliedLf) == sizeof(appliedLf))
    {
        g_state.fontName = appliedLf.lfFaceName;
        g_state.fontWeight = appliedLf.lfWeight;
        g_state.fontItalic = (appliedLf.lfItalic != 0);
        g_state.fontUnderline = (appliedLf.lfUnderline != 0);
    }

    SendMessageW(g_hwndEditor, WM_SETFONT, reinterpret_cast<WPARAM>(g_state.hFont), FALSE);

    if (!g_state.largeFileMode)
    {
        COLORREF textColor = ThemeColorEditorText(IsDarkMode());
        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_COLOR | CFM_FACE | CFM_SIZE | CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_CHARSET;
        cf.dwEffects = 0;
        if (g_state.fontWeight >= FW_BOLD)
            cf.dwEffects |= CFE_BOLD;
        if (g_state.fontItalic)
            cf.dwEffects |= CFE_ITALIC;
        if (g_state.fontUnderline)
            cf.dwEffects |= CFE_UNDERLINE;
        cf.yHeight = std::max(1, size) * 20;
        cf.bCharSet = appliedLf.lfCharSet != 0 ? static_cast<BYTE>(appliedLf.lfCharSet) : DEFAULT_CHARSET;
        wcsncpy_s(cf.szFaceName, g_state.fontName.c_str(), _TRUNCATE);
        cf.crTextColor = textColor;

        SendMessageW(g_hwndEditor, EM_SETCHARFORMAT, SCF_DEFAULT, reinterpret_cast<LPARAM>(&cf));

        CHARRANGE savedSelection{};
        SendMessageW(g_hwndEditor, EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&savedSelection));
        SendMessageW(g_hwndEditor, EM_SETSEL, 0, -1);
        SendMessageW(g_hwndEditor, EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&cf));
        SendMessageW(g_hwndEditor, EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&savedSelection));
    }

    RedrawWindow(g_hwndEditor, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}

void ApplyZoom()
{
    ApplyFont();
}

bool ScrollEditorFromMouseWheel(HWND hwndEditor, WPARAM wParam)
{
    if (!hwndEditor)
        return false;

    const int delta = GET_WHEEL_DELTA_WPARAM(wParam);
    UINT scrollLines = 3;
    SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &scrollLines, 0);
    if (scrollLines == 0)
        return true;

    if (scrollLines == static_cast<UINT>(WHEEL_PAGESCROLL))
    {
        SendMessageW(hwndEditor, WM_VSCROLL, (delta > 0) ? SB_PAGEUP : SB_PAGEDOWN, 0);
        return true;
    }

    g_verticalWheelDeltaRemainder += delta;
    const int wheelSteps = g_verticalWheelDeltaRemainder / WHEEL_DELTA;
    if (wheelSteps != 0)
    {
        g_verticalWheelDeltaRemainder -= wheelSteps * WHEEL_DELTA;
        int lineDelta = -static_cast<int>(scrollLines) * wheelSteps;
        lineDelta = std::clamp(lineDelta, -120, 120);
        SendMessageW(hwndEditor, EM_LINESCROLL, 0, lineDelta);
    }
    return true;
}

void ApplyWordWrap()
{
    std::wstring text = GetEditorText();
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    DestroyWindow(g_hwndEditor);
    const DWORD style = BuildEditorWindowStyle();
    const wchar_t *editorClass = g_editorClassName.empty() ? MSFTEDIT_CLASS : g_editorClassName.c_str();
    g_hwndEditor = CreateWindowExW(0, editorClass, nullptr, style,
                                   0, 0, 100, 100, g_hwndMain, reinterpret_cast<HMENU>(IDC_EDITOR), GetModuleHandleW(nullptr), nullptr);
    g_origEditorProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndEditor, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(EditorSubclassProc)));
    ConfigureEditorControl(g_hwndEditor);
    if (g_hwndScrollbar)
        UI::CustomScrollbar::SetTarget(g_hwndScrollbar, g_hwndEditor);
    if (g_hwndSelectionAura)
        UI::SelectionAura::SetEditor(g_hwndSelectionAura, g_hwndEditor);
    SendMessageW(g_hwndEditor, EM_EXLIMITTEXT, 0, static_cast<LPARAM>(-1));
    SendMessageW(g_hwndEditor, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE);
    ApplyEditorViewportPadding();
    ApplyFont();
    ApplyTheme();
    SetEditorText(text);
    SendMessageW(g_hwndEditor, EM_SETSEL, start, end);
    ResizeControls();
    SetFocus(g_hwndEditor);
}

void DeleteWordBackward()
{
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    if (start != end)
    {
        SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
        return;
    }
    if (start == 0)
        return;
    std::wstring text = GetEditorText();
    size_t pos = start;
    while (pos > 0 && iswspace(text[pos - 1]))
        --pos;
    while (pos > 0 && !iswspace(text[pos - 1]))
        --pos;
    SendMessageW(g_hwndEditor, EM_SETSEL, pos, start);
    SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
}

void DeleteWordForward()
{
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    if (start != end)
    {
        SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
        return;
    }
    std::wstring text = GetEditorText();
    size_t len = text.size();
    size_t pos = start;
    while (pos < len && !iswspace(text[pos]))
        ++pos;
    while (pos < len && iswspace(text[pos]))
        ++pos;
    SendMessageW(g_hwndEditor, EM_SETSEL, start, pos);
    SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
}

LRESULT CALLBACK EditorSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_ERASEBKGND:
        if (g_state.background.enabled && g_bgImage && !g_state.largeFileMode)
        {
            UpdateBackgroundBitmap(hwnd);
            if (g_bgBitmap)
            {
                HDC hdc = reinterpret_cast<HDC>(wParam);
                if (!hdc)
                    break;
                RECT rc;
                GetClientRect(hwnd, &rc);
                HDC hdcMem = CreateCompatibleDC(hdc);
                if (hdcMem)
                {
                    HBITMAP hOldBmp = reinterpret_cast<HBITMAP>(SelectObject(hdcMem, g_bgBitmap));
                    if (hOldBmp && reinterpret_cast<HGDIOBJ>(hOldBmp) != HGDI_ERROR)
                    {
                        BitBlt(hdc, 0, 0, rc.right, rc.bottom, hdcMem, 0, 0, SRCCOPY);
                        SelectObject(hdcMem, hOldBmp);
                        DeleteDC(hdcMem);
                        return 1;
                    }
                    DeleteDC(hdcMem);
                }
            }
        }
        break;
    case WM_SIZE:
        EnsureNativeScrollbarsHidden(hwnd);
        ApplyEditorViewportPadding();
        if (g_state.background.enabled && g_bgImage && g_bgBitmap && !g_state.largeFileMode)
        {
            DeleteObject(g_bgBitmap);
            g_bgBitmap = nullptr;
        }
        break;
    case WM_VSCROLL:
        if (g_hwndScrollbar) InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
        break;
    case WM_CHAR:
    {
        if (wParam == VK_RETURN && g_suppressNextReturnChar)
        {
            g_suppressNextReturnChar = false;
            return 0;
        }
        if (wParam == 3)
            break;
        if (wParam == 22)
            break;
        if (wParam == 24)
            break;
        if (wParam == 26)
            break;
        if (wParam == 25)
            break;
        if (wParam == 127 || wParam == 8)
        {
            if (GetKeyState(VK_CONTROL) & 0x8000)
                return 0;
        }
        break;
    }
    case WM_KEYDOWN:
        if (HandleTabManagementShortcut(hwnd, wParam))
            return 0;
        if (HandleInlineFormattingShortcut(hwnd, wParam))
            return 0;
        if (HandleAutoListEnterKeyDown(hwnd, wParam, lParam))
            return 0;
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            if (wParam == VK_BACK)
            {
                DeleteWordBackward();
                return 0;
            }
            if (wParam == VK_DELETE)
            {
                DeleteWordForward();
                return 0;
            }
        }
        break;
    case WM_MOUSEWHEEL:
    {
        if (LOWORD(wParam) & MK_SHIFT)
        {
            int delta = GET_WHEEL_DELTA_WPARAM(wParam);
            UINT scrollLines = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &scrollLines, 0);
            if (scrollLines == (UINT)WHEEL_PAGESCROLL)
            {
                SendMessageW(hwnd, WM_HSCROLL, (delta > 0) ? SB_PAGELEFT : SB_PAGERIGHT, 0);
            }
            else
            {
                for (UINT i = 0; i < scrollLines; ++i)
                    SendMessageW(hwnd, WM_HSCROLL, (delta > 0) ? SB_LINELEFT : SB_LINERIGHT, 0);
            }
            return 0;
        }
        ScrollEditorFromMouseWheel(hwnd, wParam);
        if (g_hwndScrollbar) InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
        return 0;
    }
    case WM_MOUSEHWHEEL:
    {
        int delta = GET_WHEEL_DELTA_WPARAM(wParam);
        UINT scrollChars = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLCHARS, 0, &scrollChars, 0);
        if (delta != 0)
        {
            for (UINT i = 0; i < scrollChars; ++i)
                SendMessageW(hwnd, WM_HSCROLL, (delta > 0) ? SB_LINERIGHT : SB_LINELEFT, 0);
            return 0;
        }
        break;
    }
    }
    return CallWindowProcW(g_origEditorProc, hwnd, msg, wParam, lParam);
}
