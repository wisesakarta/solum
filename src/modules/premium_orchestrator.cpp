/*
  Solum

  Premium header render orchestration extracted from main.cpp.
*/

#include "premium_orchestrator.h"

#include "graphics_engine.h"
#include "premium_header.h"
#include "theme.h"
#include "design_system.h"
#include "core/globals.h"

#include <algorithm>

namespace PremiumOrchestrator
{
namespace
{
Graphics::Engine g_graphics;
Premium::Header g_header;
bool g_enabled = false;
bool g_active = false;
}

static bool GetHeaderRectClient(RECT &headerRect)
{
    headerRect = {0, 0, 0, 0};
    if (!g_enabled || !g_active || !g_hwndMain || !g_hwndCommandBar)
        return false;

    RECT rcCommandBar{};
    if (!GetWindowRect(g_hwndCommandBar, &rcCommandBar))
        return false;

    POINT topLeft{rcCommandBar.left, rcCommandBar.top};
    POINT bottomRight{rcCommandBar.right, rcCommandBar.bottom};
    ScreenToClient(g_hwndMain, &topLeft);
    ScreenToClient(g_hwndMain, &bottomRight);

    headerRect.left = topLeft.x;
    headerRect.top = topLeft.y;
    headerRect.right = bottomRight.x;
    headerRect.bottom = bottomRight.y;
    return (headerRect.right > headerRect.left) && (headerRect.bottom > headerRect.top);
}

static void Activate(HWND hwndMain)
{
    if (!g_enabled || g_active || !hwndMain || !g_hwndCommandBar)
        return;

    if (!g_graphics.Initialize(hwndMain))
        return;
    if (!g_header.Initialize(&g_graphics))
    {
        g_graphics.Release();
        return;
    }

    g_header.StartReveal();
    SetTimer(hwndMain, kTimerId, 16, nullptr);
    g_active = true;
}

static void Deactivate(HWND hwndMain)
{
    if (hwndMain)
        KillTimer(hwndMain, kTimerId);
    g_graphics.Release();
    g_active = false;
}

void Initialize(HWND hwndMain)
{
    g_enabled = g_state.premiumHeaderEnabled;
    if (g_enabled)
        Activate(hwndMain);
}

void Shutdown(HWND hwndMain)
{
    Deactivate(hwndMain);
}

void SetEnabled(HWND hwndMain, bool enabled)
{
    if (g_enabled == enabled)
        return;

    g_enabled = enabled;
    if (!g_enabled)
    {
        Deactivate(hwndMain);
        return;
    }
    Activate(hwndMain);
}

bool IsEnabled()
{
    return g_enabled;
}

void OnPaint(HWND hwndMain, const RECT &paintRect)
{
    if (!g_enabled || !g_active || !hwndMain || !g_graphics.GetRenderTarget())
        return;

    RECT headerRect{};
    RECT drawRect{};
    if (!GetHeaderRectClient(headerRect))
        return;
    if (!IntersectRect(&drawRect, &paintRect, &headerRect))
        return;

    auto *rt = g_graphics.GetRenderTarget();
    g_graphics.BeginDraw();

    const D2D1_RECT_F clipRect = D2D1::RectF(
        static_cast<FLOAT>(drawRect.left),
        static_cast<FLOAT>(drawRect.top),
        static_cast<FLOAT>(drawRect.right),
        static_cast<FLOAT>(drawRect.bottom));
    rt->PushAxisAlignedClip(clipRect, D2D1_ANTIALIAS_MODE_ALIASED);

    const bool dark = IsDarkMode();
    ID2D1SolidColorBrush *headerBrush = nullptr;
    const D2D1_COLOR_F headerColor = dark
                                         ? D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f)
                                         : D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    if (SUCCEEDED(rt->CreateSolidColorBrush(headerColor, &headerBrush)) && headerBrush)
    {
        rt->FillRectangle(
            D2D1::RectF(
                static_cast<FLOAT>(headerRect.left),
                static_cast<FLOAT>(headerRect.top),
                static_cast<FLOAT>(headerRect.right),
                static_cast<FLOAT>(headerRect.bottom)),
            headerBrush);
        headerBrush->Release();
    }

    g_header.Render(headerRect);
    rt->PopAxisAlignedClip();

    const HRESULT hr = g_graphics.EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        g_graphics.Release();
        g_active = false;
        Activate(hwndMain);
    }
}

void OnSize(UINT width, UINT height)
{
    if (!g_enabled || !g_active || !g_graphics.GetRenderTarget())
        return;
    if (width == 0 || height == 0)
        return;
    g_graphics.Resize(width, height);
}

void OnDpiChanged(HWND hwndMain)
{
    if (!g_enabled || !g_active || !g_graphics.GetRenderTarget() || !hwndMain)
        return;

    RECT rcClient{};
    GetClientRect(hwndMain, &rcClient);
    const LONG clientWidth = rcClient.right - rcClient.left;
    const LONG clientHeight = rcClient.bottom - rcClient.top;
    const UINT width = (clientWidth > 0) ? static_cast<UINT>(clientWidth) : 0U;
    const UINT height = (clientHeight > 0) ? static_cast<UINT>(clientHeight) : 0U;
    if (width > 0 && height > 0)
        g_graphics.Resize(width, height);
}

bool OnTimer(HWND hwndMain, WPARAM timerId)
{
    if (timerId != kTimerId)
        return false;
    if (!g_enabled || !g_active || !hwndMain)
        return true;

    g_header.Update(0.016f);
    RECT headerRect{};
    if (GetHeaderRectClient(headerRect))
        InvalidateRect(hwndMain, &headerRect, FALSE);
    return true;
}
} // namespace PremiumOrchestrator

