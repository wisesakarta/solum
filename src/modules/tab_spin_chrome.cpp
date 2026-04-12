/*
  Technical Standard

  Tab overflow spin-button custom chrome implementation.
*/

#include "tab_spin_chrome.h"

#include <windowsx.h>
#include "theme.h"

#include <algorithm>

namespace
{
WNDPROC g_origTabSpinProc = nullptr;
HWND g_hwndTabSpin = nullptr;
bool g_tabSpinHoverLeft = false;
bool g_tabSpinHoverRight = false;
bool g_tabSpinPressLeft = false;
bool g_tabSpinPressRight = false;
bool g_trackingTabSpinMouse = false;

void FillSolidRectDc(HDC hdc, const RECT &rc, COLORREF color)
{
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(DC_BRUSH));
    const COLORREF oldBrushColor = SetDCBrushColor(hdc, color);
    FillRect(hdc, &rc, reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
    SetDCBrushColor(hdc, oldBrushColor);
    SelectObject(hdc, oldBrush);
}

LRESULT CALLBACK TabSpinSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);

        RECT rc{};
        GetClientRect(hwnd, &rc);

        const bool dark = IsDarkMode();
        const TabPaintPalette palette = GetTabPaintPalette(dark);

        FillSolidRectDc(hdc, rc, palette.stripBg);

        const int halfW = (rc.right - rc.left) / 2;
        RECT rcLeft = rc;
        rcLeft.right = rc.left + halfW;
        RECT rcRight = rc;
        rcRight.left = rcLeft.right;

        auto drawSpinButton = [&](const RECT &r, bool hover, bool press, bool isLeft)
        {
            COLORREF bg = palette.stripBg;
            if (press)
                bg = palette.inactiveBg;
            else if (hover)
                bg = palette.hoverBg;

            FillSolidRectDc(hdc, r, bg);

            const int cx = r.left + (r.right - r.left) / 2;
            const int cy = r.top + (r.bottom - r.top) / 2;

            POINT pts[3]{};
            if (isLeft)
            {
                pts[0] = {cx + 1, cy - 4};
                pts[1] = {cx - 2, cy};
                pts[2] = {cx + 1, cy + 4};
            }
            else
            {
                pts[0] = {cx - 2, cy - 4};
                pts[1] = {cx + 1, cy};
                pts[2] = {cx - 2, cy + 4};
            }

            HPEN hPen = CreatePen(PS_SOLID, 2, palette.textColor);
            HGDIOBJ oldPen = SelectObject(hdc, hPen);

            Polyline(hdc, pts, 3);

            SelectObject(hdc, oldPen);
            DeleteObject(hPen);
        };

        drawSpinButton(rcLeft, g_tabSpinHoverLeft, g_tabSpinPressLeft, true);
        drawSpinButton(rcRight, g_tabSpinHoverRight, g_tabSpinPressRight, false);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_MOUSEMOVE:
    {
        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const int halfW = (rc.right - rc.left) / 2;

        const bool hLeft = (pt.x < rc.left + halfW);
        const bool hRight = (pt.x >= rc.left + halfW);

        if (hLeft != g_tabSpinHoverLeft || hRight != g_tabSpinHoverRight)
        {
            g_tabSpinHoverLeft = hLeft;
            g_tabSpinHoverRight = hRight;
            InvalidateRect(hwnd, nullptr, FALSE);
        }

        if (!g_trackingTabSpinMouse)
        {
            TRACKMOUSEEVENT tme{};
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hwnd;
            TrackMouseEvent(&tme);
            g_trackingTabSpinMouse = true;
        }
        break;
    }
    case WM_MOUSELEAVE:
    {
        g_trackingTabSpinMouse = false;
        g_tabSpinHoverLeft = false;
        g_tabSpinHoverRight = false;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    }
    case WM_LBUTTONDOWN:
    {
        POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const int halfW = (rc.right - rc.left) / 2;

        if (pt.x < rc.left + halfW)
            g_tabSpinPressLeft = true;
        else
            g_tabSpinPressRight = true;

        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    }
    case WM_LBUTTONUP:
    {
        g_tabSpinPressLeft = false;
        g_tabSpinPressRight = false;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    }
    case WM_ERASEBKGND:
        return 1;
    }

    if (g_origTabSpinProc)
        return CallWindowProcW(g_origTabSpinProc, hwnd, msg, wParam, lParam);
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}
} // namespace

void TabSpinAttachIfNeeded(HWND hwndTabs)
{
    if (!hwndTabs)
        return;

    HWND hSpin = FindWindowExW(hwndTabs, nullptr, L"msctls_updown32", nullptr);
    if (hSpin && hSpin != g_hwndTabSpin)
    {
        if (g_hwndTabSpin && g_origTabSpinProc)
            SetWindowLongPtrW(g_hwndTabSpin, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(g_origTabSpinProc));

        g_hwndTabSpin = hSpin;
        g_origTabSpinProc = reinterpret_cast<WNDPROC>(
            SetWindowLongPtrW(hSpin, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(TabSpinSubclassProc)));
    }
    else if (!hSpin && g_hwndTabSpin)
    {
        TabSpinDetach();
    }
}

void TabSpinDetach()
{
    if (g_hwndTabSpin && g_origTabSpinProc)
        SetWindowLongPtrW(g_hwndTabSpin, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(g_origTabSpinProc));

    g_hwndTabSpin = nullptr;
    g_origTabSpinProc = nullptr;
    g_tabSpinHoverLeft = false;
    g_tabSpinHoverRight = false;
    g_tabSpinPressLeft = false;
    g_tabSpinPressRight = false;
    g_trackingTabSpinMouse = false;
}
