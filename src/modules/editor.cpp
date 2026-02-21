/*
  Saka Studio & Engineering

  Editor control functions for text manipulation, font rendering, and zoom control.
  Handles RichEdit control subclassing, word wrap, and cursor position tracking.
*/

#include "editor.h"
#include "core/types.h"
#include "core/globals.h"
#include "theme.h"
#include "ui.h"
#include "background.h"
#include "resource.h"
#include <richedit.h>
#include <algorithm>

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

static bool IsEditorWrapEnabled()
{
    return g_state.wordWrap && !g_state.largeFileMode;
}

static int ScaleEditorPx(int px)
{
    HWND ref = g_hwndEditor ? g_hwndEditor : g_hwndMain;
    if (!ref)
        return px;

    HDC hdc = GetDC(ref);
    if (!hdc)
        return px;

    const int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(ref, hdc);
    return MulDiv(px, dpi > 0 ? dpi : 96, 96);
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

std::pair<int, int> GetCursorPos()
{
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    int line = static_cast<int>(SendMessageW(g_hwndEditor, EM_EXLINEFROMCHAR, 0, start));
    int lineIndex = static_cast<int>(SendMessageW(g_hwndEditor, EM_LINEINDEX, static_cast<WPARAM>(line), 0));
    int col = static_cast<int>(start) - lineIndex;
    return {line + 1, col + 1};
}

void ConfigureEditorControl(HWND hwnd)
{
    if (!hwnd)
        return;

    // Keep RichEdit in plain-text mode and disable URL auto-detection for security/perf.
    SendMessageW(hwnd, EM_SETTEXTMODE, TM_PLAINTEXT | TM_MULTILEVELUNDO | TM_MULTICODEPAGE, 0);
    SendMessageW(hwnd, EM_AUTOURLDETECT, FALSE, 0);
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

void ApplyFont()
{
    if (g_state.hFont)
    {
        DeleteObject(g_state.hFont);
        g_state.hFont = nullptr;
    }
    int size = g_state.fontSize * g_state.zoomLevel / 100;
    size = (size < 8) ? 8 : (size > 500) ? 500
                                         : size;
    HDC hdc = GetDC(g_hwndMain);
    int height = -MulDiv(size, GetDeviceCaps(hdc, LOGPIXELSY), 72);
    ReleaseDC(g_hwndMain, hdc);
    g_state.hFont = CreateFontW(height, 0, 0, 0, g_state.fontWeight, g_state.fontItalic, g_state.fontUnderline, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                DEFAULT_PITCH | FF_DONTCARE, g_state.fontName.c_str());
    SendMessageW(g_hwndEditor, WM_SETFONT, reinterpret_cast<WPARAM>(g_state.hFont), TRUE);
    COLORREF textColor = IsDarkMode() ? RGB(255, 255, 255) : GetSysColor(COLOR_WINDOWTEXT);
    CHARFORMAT2W cf = {};
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_COLOR;
    cf.crTextColor = textColor;
    SendMessageW(g_hwndEditor, EM_SETCHARFORMAT, SCF_ALL, reinterpret_cast<LPARAM>(&cf));
    SendMessageW(g_hwndEditor, EM_SETCHARFORMAT, SCF_DEFAULT, reinterpret_cast<LPARAM>(&cf));
}

void ApplyZoom()
{
    ApplyFont();
}

void ApplyWordWrap()
{
    std::wstring text = GetEditorText();
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    DestroyWindow(g_hwndEditor);
    DWORD style = WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN | ES_NOHIDESEL;
    if (!IsEditorWrapEnabled())
        style |= WS_HSCROLL | ES_AUTOHSCROLL;
    const wchar_t *editorClass = g_editorClassName.empty() ? MSFTEDIT_CLASS : g_editorClassName.c_str();
    g_hwndEditor = CreateWindowExW(0, editorClass, nullptr, style,
                                   0, 0, 100, 100, g_hwndMain, reinterpret_cast<HMENU>(IDC_EDITOR), GetModuleHandleW(nullptr), nullptr);
    g_origEditorProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(g_hwndEditor, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(EditorSubclassProc)));
    ConfigureEditorControl(g_hwndEditor);
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
                RECT rc;
                GetClientRect(hwnd, &rc);
                HDC hdcMem = CreateCompatibleDC(hdc);
                HBITMAP hOldBmp = reinterpret_cast<HBITMAP>(SelectObject(hdcMem, g_bgBitmap));
                BitBlt(hdc, 0, 0, rc.right, rc.bottom, hdcMem, 0, 0, SRCCOPY);
                SelectObject(hdcMem, hOldBmp);
                DeleteDC(hdcMem);
                return 1;
            }
        }
        break;
    case WM_SIZE:
        ApplyEditorViewportPadding();
        if (g_state.background.enabled && g_bgImage && g_bgBitmap && !g_state.largeFileMode)
        {
            DeleteObject(g_bgBitmap);
            g_bgBitmap = nullptr;
        }
        break;
    case WM_CHAR:
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
        if (g_state.background.enabled && g_bgImage && !g_state.largeFileMode)
        {
            LRESULT result = CallWindowProcW(g_origEditorProc, hwnd, msg, wParam, lParam);
            InvalidateRect(hwnd, nullptr, TRUE);
            return result;
        }
        break;
    case WM_KEYDOWN:
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
        if (g_state.background.enabled && g_bgImage && !g_state.largeFileMode && (wParam == VK_BACK || wParam == VK_DELETE))
        {
            LRESULT result = CallWindowProcW(g_origEditorProc, hwnd, msg, wParam, lParam);
            InvalidateRect(hwnd, nullptr, TRUE);
            return result;
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
        break;
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
