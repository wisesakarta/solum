/*
  Technical Standard

  Dialog box implementations for find, replace, goto, font selection, and more.
  Provides modeless and modal dialog creation with proper event handling.
*/

#include "dialog.h"
#include "core/globals.h"
#include "editor.h"
#include "ui.h"
#include "settings.h"
#include "theme.h"
#include "lang/lang.h"
#include "background.h"
#include "design_system.h"
#include "resource.h"
#include "tab_layout.h"
#include <commdlg.h>
#include <richedit.h>
#include <algorithm>
#include <cwctype>

namespace
{
constexpr COLORREF kDialogDarkBg = DesignSystem::Color::kDarkBg;
constexpr COLORREF kDialogDarkEditBg = DesignSystem::Color::kDarkBg;
constexpr COLORREF kDialogDarkText = DesignSystem::Color::kDarkInk;

INT_PTR HandleDialogPaint(HWND hDlg);
INT_PTR HandleDialogCtlColor(UINT msg, WPARAM wParam);

HFONT DialogUiFont()
{
    HFONT hFont = TabGetRegularFont();
    if (!hFont)
        hFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    return hFont;
}

int ScaleDialogPx(int px)
{
    return DesignSystem::ScalePx(px, g_hwndMain);
}

int MeasureDialogTextWidth(const std::wstring &text)
{
    HWND ref = g_hwndMain ? g_hwndMain : GetDesktopWindow();
    HDC hdc = GetDC(ref);
    if (!hdc)
        return 0;

    const HFONT hFont = DialogUiFont();
    HGDIOBJ oldFont = SelectObject(hdc, hFont);
    SIZE size{};
    GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size);
    SelectObject(hdc, oldFont);
    ReleaseDC(ref, hdc);
    return size.cx;
}

UINT_PTR CALLBACK FontDialogHookProc(HWND hDlg, UINT msg, WPARAM, LPARAM)
{
    if (msg != WM_INITDIALOG)
        return FALSE;

    HWND hRoot = GetAncestor(hDlg, GA_ROOT);
    if (!hRoot)
        hRoot = hDlg;

    SetTitleBarDark(hRoot, IsDarkMode() ? TRUE : FALSE);
    ApplyThemeToWindowTree(hRoot);

    const HFONT hUiFont = DialogUiFont();

    for (HWND h = GetWindow(hRoot, GW_CHILD); h; h = GetWindow(h, GW_HWNDNEXT))
        SendMessageW(h, WM_SETFONT, reinterpret_cast<WPARAM>(hUiFont), TRUE);

    return FALSE;
}

void ApplyDialogTheme(HWND hDlg)
{
    if (!hDlg)
        return;
    ApplyThemeToWindowTree(hDlg);
    InvalidateRect(hDlg, nullptr, TRUE);
    UpdateWindow(hDlg);
}

INT_PTR HandleDialogPaint(HWND hDlg)
{
    PAINTSTRUCT ps{};
    HDC hdc = BeginPaint(hDlg, &ps);

    HBRUSH hBrush = nullptr;
    HBRUSH hTmpBrush = nullptr;
    if (IsDarkMode())
    {
        hBrush = g_hbrDialogDark;
        if (!hBrush)
        {
            hTmpBrush = CreateSolidBrush(kDialogDarkBg);
            hBrush = hTmpBrush;
        }
    }
    else
    {
        hBrush = GetSysColorBrush(COLOR_BTNFACE);
    }

    FillRect(hdc, &ps.rcPaint, hBrush);
    EndPaint(hDlg, &ps);
    if (hTmpBrush)
        DeleteObject(hTmpBrush);
    return TRUE;
}

INT_PTR HandleDialogCtlColor(UINT msg, WPARAM wParam)
{
    if (!IsDarkMode())
    {
        if (msg == WM_CTLCOLOREDIT || msg == WM_CTLCOLORLISTBOX)
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_WINDOW));
        return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_BTNFACE));
    }

    HDC hdc = reinterpret_cast<HDC>(wParam);
    SetTextColor(hdc, kDialogDarkText);
    if (msg == WM_CTLCOLOREDIT || msg == WM_CTLCOLORLISTBOX)
    {
        SetBkColor(hdc, kDialogDarkEditBg);
        SetBkMode(hdc, OPAQUE);
        return reinterpret_cast<INT_PTR>(g_hbrDialogEditDark ? g_hbrDialogEditDark : GetStockObject(BLACK_BRUSH));
    }

    SetBkMode(hdc, TRANSPARENT);
    SetBkColor(hdc, kDialogDarkBg);
    return reinterpret_cast<INT_PTR>(g_hbrDialogDark ? g_hbrDialogDark : GetStockObject(BLACK_BRUSH));
}

static void FillSolidRectDc(HDC hdc, const RECT &rc, COLORREF color)
{
    HBRUSH hBrush = CreateSolidBrush(color);
    FillRect(hdc, &rc, hBrush);
    DeleteObject(hBrush);
}

INT_PTR HandleAboutPaint(HWND hWnd)
{
    PAINTSTRUCT ps{};
    HDC hdc = BeginPaint(hWnd, &ps);

    RECT rcClient{};
    GetClientRect(hWnd, &rcClient);

    const bool dark = IsDarkMode();
    const TabPaintPalette palette = GetTabPaintPalette(dark);
    
    FillSolidRectDc(hdc, rcClient, palette.stripBg);

    if (EnsureBackgroundGraphicsReady())
    {
        Gdiplus::Graphics graphics(hdc);
        graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);

        Gdiplus::Image* logo = LoadImageFromResource(GetModuleHandleW(nullptr), IDR_LOGO_WORDMARK, L"RCDATA");
        if (logo)
        {
            float imgW = static_cast<float>(logo->GetWidth());
            float imgH = static_cast<float>(logo->GetHeight());
            
            float targetW = static_cast<float>(ScaleDialogPx(200));
            float targetH = (imgH / imgW) * targetW;
            
            float x = (static_cast<float>(rcClient.right) - targetW) / 2.0f;
            float y = static_cast<float>(ScaleDialogPx(40));

            graphics.DrawImage(logo, x, y, targetW, targetH);
            delete logo;
        }
    }


    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, palette.textColor);

    const HFONT hFontReg = DialogUiFont();
    HGDIOBJ oldFont = SelectObject(hdc, hFontReg);

    std::wstring verText = L"v" + std::wstring(APP_VERSION);
    RECT rcVer = { 0, ScaleDialogPx(130), rcClient.right, ScaleDialogPx(155) };
    DrawTextW(hdc, verText.c_str(), -1, &rcVer, DT_CENTER | DT_SINGLELINE | DT_NOPREFIX);

    std::wstring copyText = L"Crafted with discipline by wisesakarta";
    RECT rcCopy = { 0, ScaleDialogPx(160), rcClient.right, ScaleDialogPx(185) };
    DrawTextW(hdc, copyText.c_str(), -1, &rcCopy, DT_CENTER | DT_SINGLELINE | DT_NOPREFIX);

    std::wstring subText = L"Clarity. Function. Detail.";
    RECT rcSub = { 0, ScaleDialogPx(180), rcClient.right, ScaleDialogPx(205) };
    DrawTextW(hdc, subText.c_str(), -1, &rcSub, DT_CENTER | DT_SINGLELINE | DT_NOPREFIX);

    SelectObject(hdc, oldFont);
    EndPaint(hWnd, &ps);
    return 0;
}

LRESULT CALLBACK AboutWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
        SetTitleBarDark(hWnd, IsDarkMode());
        return 0;
    case WM_PAINT:
        HandleAboutPaint(hWnd);
        return 0;
    case WM_LBUTTONDOWN:
    case WM_KEYDOWN:
        DestroyWindow(hWnd);
        return 0;
    case WM_CLOSE:
        DestroyWindow(hWnd);
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

}

void DoFind(bool forward)
{
    if (g_state.findText.empty())
        return;
    std::wstring text = GetEditorText();
    DWORD start = 0, end = 0;
    SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
    std::wstring textLower = text;
    std::transform(textLower.begin(), textLower.end(), textLower.begin(), towlower);
    std::wstring findLower = g_state.findText;
    std::transform(findLower.begin(), findLower.end(), findLower.begin(), towlower);
    size_t pos = std::wstring::npos;
    if (forward)
    {
        pos = textLower.find(findLower, end);
        if (pos == std::wstring::npos)
            pos = textLower.find(findLower);
    }
    else
    {
        if (start > 0)
            pos = textLower.rfind(findLower, start - 1);
        if (pos == std::wstring::npos)
            pos = textLower.rfind(findLower);
    }
    if (pos != std::wstring::npos)
    {
        SendMessageW(g_hwndEditor, EM_SETSEL, pos, pos + g_state.findText.size());
        SendMessageW(g_hwndEditor, EM_SCROLLCARET, 0, 0);
    }
    else
    {
        const auto &lang = GetLangStrings();
        MessageBoxW(g_hwndMain, (lang.msgCannotFind + g_state.findText + L"\"").c_str(), lang.appName.c_str(), MB_ICONINFORMATION);
    }
}

INT_PTR CALLBACK FindDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_INITDIALOG:
        SetWindowTextW(GetDlgItem(hDlg, 1001), g_state.findText.c_str());
        if (GetDlgItem(hDlg, 1002))
            SetWindowTextW(GetDlgItem(hDlg, 1002), g_state.replaceText.c_str());
        InvalidateRect(hDlg, nullptr, FALSE);
        return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case 1:
        {
            wchar_t buf[256] = {0};
            GetWindowTextW(GetDlgItem(hDlg, 1001), buf, 256);
            g_state.findText = buf;
            DoFind(true);
            return TRUE;
        }
        case 2:
            DestroyWindow(hDlg);
            g_hwndFindDlg = nullptr;
            SetFocus(g_hwndEditor);
            return TRUE;
        case 3:
        {
            wchar_t buf[256] = {0};
            GetWindowTextW(GetDlgItem(hDlg, 1001), buf, 256);
            g_state.findText = buf;
            GetWindowTextW(GetDlgItem(hDlg, 1002), buf, 256);
            g_state.replaceText = buf;
            if (g_state.findText.empty())
                return TRUE;
            DWORD start = 0, end = 0;
            SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
            if (start != end)
            {
                std::wstring text = GetEditorText();
                std::wstring sel = text.substr(start, end - start);
                std::transform(sel.begin(), sel.end(), sel.begin(), towlower);
                std::wstring findLower = g_state.findText;
                std::transform(findLower.begin(), findLower.end(), findLower.begin(), towlower);
                if (sel == findLower)
                    SendMessageW(g_hwndEditor, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(g_state.replaceText.c_str()));
            }
            DoFind(true);
            return TRUE;
        }
        case 4:
        {
            wchar_t buf[256] = {0};
            GetWindowTextW(GetDlgItem(hDlg, 1001), buf, 256);
            g_state.findText = buf;
            GetWindowTextW(GetDlgItem(hDlg, 1002), buf, 256);
            g_state.replaceText = buf;
            if (g_state.findText.empty())
                return TRUE;
            std::wstring text = GetEditorText();
            std::wstring findLower = g_state.findText;
            std::transform(findLower.begin(), findLower.end(), findLower.begin(), towlower);
            std::wstring lower = text;
            std::transform(lower.begin(), lower.end(), lower.begin(), towlower);
            std::wstring newText;
            size_t lastPos = 0, pos = 0;
            while ((pos = lower.find(findLower, lastPos)) != std::wstring::npos)
            {
                newText += text.substr(lastPos, pos - lastPos);
                newText += g_state.replaceText;
                lastPos = pos + g_state.findText.size();
            }
            newText += text.substr(lastPos);
            if (newText != text)
            {
                SetEditorText(newText);
                g_state.modified = true;
                UpdateTitle();
            }
            return TRUE;
        }
        }
        break;
    case WM_PAINT:
        return HandleDialogPaint(hDlg);
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORDLG:
        return HandleDialogCtlColor(msg, wParam);
    case WM_CLOSE:
        DestroyWindow(hDlg);
        g_hwndFindDlg = nullptr;
        SetFocus(g_hwndEditor);
        return TRUE;
    case WM_DESTROY:
        g_hwndFindDlg = nullptr;
        return TRUE;
    }
    return DefDlgProcW(hDlg, msg, wParam, lParam);
}

void EditFind()
{
    if (g_hwndFindDlg)
    {
        SetFocus(g_hwndFindDlg);
        return;
    }
    const auto &lang = GetLangStrings();
    g_hwndFindDlg = CreateWindowExW(WS_EX_DLGMODALFRAME, L"#32770", lang.dialogFind.c_str(),
                                    WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE, 100, 100, 420, 120,
                                    g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (g_hwndFindDlg)
    {
        HFONT hFont = DialogUiFont();
        CreateWindowExW(0, L"STATIC", lang.dialogFindLabel.c_str(), WS_CHILD | WS_VISIBLE, 10, 12, 45, 16, g_hwndFindDlg, nullptr, nullptr, nullptr);
        CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", g_state.findText.c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 60, 10, 230, 20, g_hwndFindDlg, reinterpret_cast<HMENU>(1001), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogFindNext.c_str(), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 300, 10, 100, 22, g_hwndFindDlg, reinterpret_cast<HMENU>(1), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogClose.c_str(), WS_CHILD | WS_VISIBLE, 300, 38, 100, 22, g_hwndFindDlg, reinterpret_cast<HMENU>(2), nullptr, nullptr);
        for (HWND h = GetWindow(g_hwndFindDlg, GW_CHILD); h; h = GetWindow(h, GW_HWNDNEXT))
            SendMessageW(h, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        SetWindowLongPtrW(g_hwndFindDlg, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(FindDlgProc));
        ApplyDialogTheme(g_hwndFindDlg);
    }
}

void EditFindNext()
{
    if (!g_state.findText.empty())
        DoFind(true);
}

void EditFindPrev()
{
    if (!g_state.findText.empty())
        DoFind(false);
}

void EditReplace()
{
    if (g_hwndFindDlg)
    {
        SetFocus(g_hwndFindDlg);
        return;
    }
    const auto &lang = GetLangStrings();
    g_hwndFindDlg = CreateWindowExW(WS_EX_DLGMODALFRAME, L"#32770", lang.dialogFindReplace.c_str(),
                                    WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE, 100, 100, 420, 175,
                                    g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (g_hwndFindDlg)
    {
        HFONT hFont = DialogUiFont();
        CreateWindowExW(0, L"STATIC", lang.dialogFindLabel.c_str(), WS_CHILD | WS_VISIBLE, 10, 12, 45, 16, g_hwndFindDlg, nullptr, nullptr, nullptr);
        CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", g_state.findText.c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 60, 10, 230, 20, g_hwndFindDlg, reinterpret_cast<HMENU>(1001), nullptr, nullptr);
        CreateWindowExW(0, L"STATIC", lang.dialogReplaceLabel.c_str(), WS_CHILD | WS_VISIBLE, 10, 40, 50, 16, g_hwndFindDlg, nullptr, nullptr, nullptr);
        CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", g_state.replaceText.c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 60, 38, 230, 20, g_hwndFindDlg, reinterpret_cast<HMENU>(1002), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogFindNext.c_str(), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 300, 10, 100, 22, g_hwndFindDlg, reinterpret_cast<HMENU>(1), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogReplace.c_str(), WS_CHILD | WS_VISIBLE, 300, 38, 100, 22, g_hwndFindDlg, reinterpret_cast<HMENU>(3), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogReplaceAll.c_str(), WS_CHILD | WS_VISIBLE, 300, 66, 100, 22, g_hwndFindDlg, reinterpret_cast<HMENU>(4), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogClose.c_str(), WS_CHILD | WS_VISIBLE, 300, 94, 100, 22, g_hwndFindDlg, reinterpret_cast<HMENU>(2), nullptr, nullptr);
        for (HWND h = GetWindow(g_hwndFindDlg, GW_CHILD); h; h = GetWindow(h, GW_HWNDNEXT))
            SendMessageW(h, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        SetWindowLongPtrW(g_hwndFindDlg, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(FindDlgProc));
        ApplyDialogTheme(g_hwndFindDlg);
    }
}

INT_PTR CALLBACK GotoDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_PAINT:
        return HandleDialogPaint(hDlg);
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORDLG:
        return HandleDialogCtlColor(msg, wParam);
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            wchar_t buf[32];
            GetWindowTextW(GetDlgItem(hDlg, 1001), buf, 32);
            int line = _wtoi(buf);
            if (line > 0)
            {
                LRESULT charIndex = SendMessageW(g_hwndEditor, EM_LINEINDEX, static_cast<WPARAM>(line) - 1, 0);
                if (charIndex != -1)
                {
                    SendMessageW(g_hwndEditor, EM_SETSEL, charIndex, charIndex);
                    SendMessageW(g_hwndEditor, EM_SCROLLCARET, 0, 0);
                    SetFocus(g_hwndEditor);
                    DestroyWindow(hDlg);
                }
                else
                {
                    const auto &lang = GetLangStrings();
                    MessageBoxW(hDlg, lang.msgLineNumberOutOfRange.c_str(), (lang.appName + L" - " + lang.dialogGoTo).c_str(), MB_OK | MB_ICONWARNING);
                }
            }
            return TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL)
        {
            DestroyWindow(hDlg);
            return TRUE;
        }
        break;
    case WM_CLOSE:
        DestroyWindow(hDlg);
        return TRUE;
    }
    return DefDlgProcW(hDlg, msg, wParam, lParam);
}

void EditGoto()
{
    const auto &lang = GetLangStrings();
    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME, L"#32770", lang.dialogGoTo.c_str(),
                                WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE, 100, 100, 250, 140,
                                g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (hDlg)
    {
        HFONT hFont = DialogUiFont();
        CreateWindowExW(0, L"STATIC", lang.dialogLineNumber.c_str(), WS_CHILD | WS_VISIBLE, 15, 15, 100, 16, hDlg, nullptr, nullptr, nullptr);

        DWORD start = 0;
        SendMessageW(g_hwndEditor, EM_GETSEL, reinterpret_cast<WPARAM>(&start), 0);
        int curLine = (int)SendMessageW(g_hwndEditor, EM_EXLINEFROMCHAR, 0, start) + 1;
        wchar_t buf[32];
        wsprintfW(buf, L"%d", curLine);

        HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", buf, WS_CHILD | WS_VISIBLE | ES_NUMBER | ES_AUTOHSCROLL, 15, 35, 210, 22, hDlg, reinterpret_cast<HMENU>(1001), nullptr, nullptr);
        SendMessageW(hEdit, EM_SETSEL, 0, -1);

        CreateWindowExW(0, L"BUTTON", lang.dialogOK.c_str(), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 60, 70, 80, 25, hDlg, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
        CreateWindowExW(0, L"BUTTON", lang.dialogCancel.c_str(), WS_CHILD | WS_VISIBLE, 145, 70, 80, 25, hDlg, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);

        for (HWND h = GetWindow(hDlg, GW_CHILD); h; h = GetWindow(h, GW_HWNDNEXT))
            SendMessageW(h, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

        SetWindowLongPtrW(hDlg, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(GotoDlgProc));
        ApplyDialogTheme(hDlg);
        SetFocus(hEdit);
    }
}

void FormatFont()
{
    LOGFONTW lf{};
    if (g_state.hFont)
        GetObjectW(g_state.hFont, sizeof(LOGFONTW), &lf);
    else
    {
        HDC hdc = GetDC(g_hwndMain);
        lf.lfHeight = -MulDiv(g_state.fontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72);
        ReleaseDC(g_hwndMain, hdc);
        wcscpy_s(lf.lfFaceName, g_state.fontName.c_str());
        lf.lfWeight = g_state.fontWeight;
        lf.lfItalic = g_state.fontItalic ? TRUE : FALSE;
        lf.lfUnderline = g_state.fontUnderline ? TRUE : FALSE;
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = CLEARTYPE_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    }

    CHOOSEFONTW cf{};
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner = g_hwndMain;
    cf.lpLogFont = &lf;
    cf.Flags = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT | CF_FORCEFONTEXIST | CF_BOTH | CF_EFFECTS | CF_ENABLEHOOK;
    cf.lpfnHook = FontDialogHookProc;

    if (ChooseFontW(&cf))
    {
        g_state.fontName = lf.lfFaceName;
        g_state.fontWeight = lf.lfWeight;
        g_state.fontItalic = (lf.lfItalic != 0);
        g_state.fontUnderline = (lf.lfUnderline != 0);
        HDC hdc2 = GetDC(g_hwndMain);
        g_state.fontSize = MulDiv(-lf.lfHeight, 72, GetDeviceCaps(hdc2, LOGPIXELSY));
        ReleaseDC(g_hwndMain, hdc2);
        ApplyFont();
        if (g_state.largeFileMode)
        {
            const int firstVisibleBefore = static_cast<int>(SendMessageW(g_hwndEditor, EM_GETFIRSTVISIBLELINE, 0, 0));
            ApplyWordWrap();
            const int firstVisibleAfter = static_cast<int>(SendMessageW(g_hwndEditor, EM_GETFIRSTVISIBLELINE, 0, 0));
            const int lineDelta = firstVisibleBefore - firstVisibleAfter;
            if (lineDelta != 0)
                SendMessageW(g_hwndEditor, EM_LINESCROLL, 0, lineDelta);
        }
        if (g_hwndScrollbar)
            InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
        SaveFontSettings();
    }
}

static INT_PTR CALLBACK TransparencyDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_PAINT:
        return HandleDialogPaint(hDlg);
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORDLG:
        return HandleDialogCtlColor(msg, wParam);
    case WM_CLOSE:
        DestroyWindow(hDlg);
        return TRUE;
    }
    return DefDlgProcW(hDlg, msg, wParam, lParam);
}

void ViewTransparency()
{
    const auto &lang = GetLangStrings();
    int pct = g_state.windowOpacity * 100 / 255;
    wchar_t buf[32];
    wsprintfW(buf, L"%d", pct);

    const int margin = ScaleDialogPx(18);
    const int gap = ScaleDialogPx(12);
    const int minLabelW = ScaleDialogPx(100);
    const int measuredLabelW = MeasureDialogTextWidth(lang.dialogOpacityLabel) + ScaleDialogPx(8);
    const int labelW = std::max(minLabelW, measuredLabelW);
    const int labelH = ScaleDialogPx(20);
    const int inputW = ScaleDialogPx(96);
    const int inputH = ScaleDialogPx(24);
    const int topPad = ScaleDialogPx(20);
    const int rowGap = ScaleDialogPx(30);
    const int buttonW = ScaleDialogPx(84);
    const int buttonH = ScaleDialogPx(30);
    const int buttonGap = ScaleDialogPx(10);
    const int bottomPad = ScaleDialogPx(20);

    const int clientW = std::max(ScaleDialogPx(360), margin + labelW + gap + inputW + margin);
    const int clientH = topPad + std::max(labelH, inputH) + rowGap + buttonH + bottomPad;

    const DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE;
    const DWORD exStyle = WS_EX_DLGMODALFRAME;
    RECT windowRect = {0, 0, clientW, clientH};
    AdjustWindowRectEx(&windowRect, style, FALSE, exStyle);
    const int windowW = windowRect.right - windowRect.left;
    const int windowH = windowRect.bottom - windowRect.top;

    int x = 300;
    int y = 300;
    RECT ownerRect{};
    if (g_hwndMain && GetWindowRect(g_hwndMain, &ownerRect))
    {
        x = ownerRect.left + ((ownerRect.right - ownerRect.left) - windowW) / 2;
        y = ownerRect.top + ((ownerRect.bottom - ownerRect.top) - windowH) / 2;
    }

    HWND hDlg = CreateWindowExW(exStyle, L"#32770", lang.dialogTransparency.c_str(),
                                style, x, y, windowW, windowH,
                                g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!hDlg)
        return;

    const int rowY = topPad;
    const int labelY = rowY + std::max(0, (inputH - labelH) / 2);
    const int inputX = margin + labelW + gap;
    const int buttonY = clientH - bottomPad - buttonH;
    const int cancelX = clientW - margin - buttonW;
    const int okX = cancelX - buttonGap - buttonW;

    HFONT hFont = DialogUiFont();
    
    CreateWindowExW(0, L"STATIC", lang.dialogOpacityLabel.c_str(), WS_CHILD | WS_VISIBLE, margin, labelY, labelW, labelH, hDlg, nullptr, nullptr, nullptr);
    HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", buf, WS_CHILD | WS_VISIBLE | ES_NUMBER | ES_AUTOHSCROLL | ES_RIGHT,
                                 inputX, rowY, inputW, inputH, hDlg, reinterpret_cast<HMENU>(1001), nullptr, nullptr);
    CreateWindowExW(0, L"BUTTON", lang.dialogOK.c_str(), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                    okX, buttonY, buttonW, buttonH, hDlg, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
    CreateWindowExW(0, L"BUTTON", lang.dialogCancel.c_str(), WS_CHILD | WS_VISIBLE,
                    cancelX, buttonY, buttonW, buttonH, hDlg, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);
    for (HWND h = GetWindow(hDlg, GW_CHILD); h; h = GetWindow(h, GW_HWNDNEXT))
        SendMessageW(h, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
    SetWindowLongPtrW(hDlg, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(TransparencyDlgProc));
    ApplyDialogTheme(hDlg);
    SetFocus(hEdit);
    MSG msg;
    int quitCode = -1;
    while (IsWindow(hDlg))
    {
        BOOL gm = GetMessageW(&msg, nullptr, 0, 0);
        if (gm == 0)
        {
            quitCode = static_cast<int>(msg.wParam);
            break;
        }
        if (gm < 0)
            break;
        if (msg.hwnd == hDlg && msg.message == WM_COMMAND)
        {
            if (LOWORD(msg.wParam) == IDOK)
            {
                GetWindowTextW(hEdit, buf, 32);
                int val = _wtoi(buf);
                val = (val < 10) ? 10 : (val > 100) ? 100
                                                    : val;
                g_state.windowOpacity = static_cast<BYTE>(val * 255 / 100);
                SetWindowLongW(g_hwndMain, GWL_EXSTYLE, GetWindowLongW(g_hwndMain, GWL_EXSTYLE) | WS_EX_LAYERED);
                SetLayeredWindowAttributes(g_hwndMain, 0, g_state.windowOpacity, LWA_ALPHA);
                SaveFontSettings();
                DestroyWindow(hDlg);
                break;
            }
            if (LOWORD(msg.wParam) == IDCANCEL)
            {
                DestroyWindow(hDlg);
                break;
            }
        }
        if (msg.hwnd == hDlg && msg.message == WM_CLOSE)
        {
            DestroyWindow(hDlg);
            break;
        }
        if (!IsDialogMessageW(hDlg, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    if (IsWindow(hDlg))
        DestroyWindow(hDlg);
    if (quitCode >= 0)
        PostQuitMessage(quitCode);
}

void HelpAbout()
{
    const auto &lang = GetLangStrings();
    const wchar_t* clsName = L"TechnicalStandardAboutBox";

    static bool clsRegistered = false;
    if (!clsRegistered)
    {
        WNDCLASSEXW wc;
        ZeroMemory(&wc, sizeof(wc));
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = AboutWndProc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.lpszClassName = clsName;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wc);
        clsRegistered = true;
    }

    int winW = ScaleDialogPx(380);
    int winH = ScaleDialogPx(260);

    RECT rcMain;
    GetWindowRect(g_hwndMain, &rcMain);
    int x = rcMain.left + (rcMain.right - rcMain.left - winW) / 2;
    int y = rcMain.top + (rcMain.bottom - rcMain.top - winH) / 2;

    HWND hAbout = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_DLGMODALFRAME, clsName, lang.menuAbout.c_str(),
                                  WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
                                  x, y, winW, winH, g_hwndMain, nullptr, GetModuleHandleW(nullptr), nullptr);
    
    if (hAbout)
    {
        EnableWindow(g_hwndMain, FALSE);
        MSG msg;
        while (IsWindow(hAbout))
        {
            if (GetMessageW(&msg, nullptr, 0, 0))
            {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
        EnableWindow(g_hwndMain, TRUE);
        SetForegroundWindow(g_hwndMain);
    }
}
