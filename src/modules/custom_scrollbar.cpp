#include "custom_scrollbar.h"
#include "design_system.h"
#include "editor.h"
#include "theme.h"
#include <commctrl.h>
#include <richedit.h>
#include <algorithm>

namespace UI {

bool CustomScrollbar::RegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"TechnicalStandardCustomScrollbar";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    return ::RegisterClassExW(&wc) != 0;
}

HWND CustomScrollbar::Create(HWND hwndParent, HWND hwndTarget)
{
    HWND hwnd = CreateWindowExW(0, L"TechnicalStandardCustomScrollbar", nullptr,
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0,
        hwndParent, nullptr, GetModuleHandle(nullptr), hwndTarget);
    return hwnd;
}

void CustomScrollbar::SetTarget(HWND hwndScrollbar, HWND hwndTarget)
{
    if (!hwndScrollbar)
        return;
    State* state = reinterpret_cast<State*>(GetWindowLongPtrW(hwndScrollbar, GWLP_USERDATA));
    if (!state)
        return;
    state->hwndTarget = hwndTarget;
    InvalidateRect(hwndScrollbar, nullptr, FALSE);
}

LRESULT CALLBACK CustomScrollbar::WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    State* state = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg) {
    case WM_CREATE: {
        State* state = new State();
        state->hwndTarget = reinterpret_cast<HWND>(reinterpret_cast<CREATESTRUCT*>(lParam)->lpCreateParams);
        state->engine.Initialize(hwnd);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
        return 0;
    }
    case WM_SIZE: {
        if (state) state->engine.Resize(LOWORD(lParam), HIWORD(lParam));
        return 0;
    }
    case WM_MOUSEMOVE: {
        if (!state) return 0;
        if (!state->isHovered) {
            state->isHovered = true;
            TRACKMOUSEEVENT tme = { sizeof(tme), TME_LEAVE, hwnd, 0 };
            TrackMouseEvent(&tme);
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        
        if (state->isDragging && state->hwndTarget) {
            RECT rc;
            GetClientRect(hwnd, &rc);
            POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
            
            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
            if (!GetScrollInfo(state->hwndTarget, SB_VERT, &si))
                return 0;
            int trackHeight = rc.bottom - rc.top;
            float thumbSize = si.nMax > si.nMin ? (float)(trackHeight * si.nPage) / (si.nMax - si.nMin + 1) : 20.0f;
            thumbSize = std::max(thumbSize, 20.0f);

            const int usableTrackHeight = std::max(1, trackHeight - static_cast<int>(thumbSize));
            float scrollRatio = static_cast<float>(pt.y - static_cast<int>(thumbSize) / 2) / static_cast<float>(usableTrackHeight);
            scrollRatio = std::clamp(scrollRatio, 0.0f, 1.0f);

            const int maxPos = std::max(si.nMin, si.nMax - static_cast<int>(si.nPage) + 1);
            int newPos = si.nMin + static_cast<int>(scrollRatio * (maxPos - si.nMin));
            newPos = std::clamp(newPos, si.nMin, maxPos);
            int lineDelta = newPos - si.nPos;
            lineDelta = std::clamp(lineDelta, -256, 256);
            if (lineDelta != 0)
            {
                const int pageStep = std::max(1, static_cast<int>(si.nPage) - 1);
                while (lineDelta >= pageStep)
                {
                    SendMessageW(state->hwndTarget, WM_VSCROLL, SB_PAGEDOWN, 0);
                    lineDelta -= pageStep;
                }
                while (lineDelta <= -pageStep)
                {
                    SendMessageW(state->hwndTarget, WM_VSCROLL, SB_PAGEUP, 0);
                    lineDelta += pageStep;
                }
                while (lineDelta > 0)
                {
                    SendMessageW(state->hwndTarget, WM_VSCROLL, SB_LINEDOWN, 0);
                    --lineDelta;
                }
                while (lineDelta < 0)
                {
                    SendMessageW(state->hwndTarget, WM_VSCROLL, SB_LINEUP, 0);
                    ++lineDelta;
                }
            }
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    }
    case WM_MOUSELEAVE:
        if (state) {
            state->isHovered = false;
            if (!state->isDragging) InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_LBUTTONDOWN:
        if (state) {
            state->isDragging = true;
            SetCapture(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_LBUTTONUP:
        if (state) {
            state->isDragging = false;
            ReleaseCapture();
            if (state->hwndTarget)
                SetFocus(state->hwndTarget);
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_MOUSEWHEEL:
    case WM_MOUSEHWHEEL:
        if (state && state->hwndTarget) {
            if (msg == WM_MOUSEWHEEL)
                ScrollEditorFromMouseWheel(state->hwndTarget, wParam);
            else
                SendMessageW(state->hwndTarget, msg, wParam, lParam);
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        }
        return 0;
    case WM_PAINT: {
        if (!state) return 0;
        PAINTSTRUCT ps;
        BeginPaint(hwnd, &ps);
        
        auto context = state->engine.GetDeviceContext();
        if (context) {
            RECT rc;
            GetClientRect(hwnd, &rc);
            
            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask = SIF_ALL;
            if (state->hwndTarget)
                GetScrollInfo(state->hwndTarget, SB_VERT, &si);
            
            float trackHeight = (float)(rc.bottom - rc.top);
            float thumbHeight = si.nMax > si.nMin ? (trackHeight * si.nPage) / (si.nMax - si.nMin + 1) : 0;
            thumbHeight = std::max(thumbHeight, 20.0f); // Min height
            
            float scrollRange = (float)(si.nMax - si.nMin - (int)si.nPage + 1);
            float thumbY = scrollRange > 0 ? (si.nPos * (trackHeight - thumbHeight)) / scrollRange : 0;

            context->BeginDraw();
            
            const bool dark = IsDarkMode();
            D2D1_COLOR_F bgColor = Graphics::Engine::ColorToD2D(dark ? DesignSystem::Color::kDarkBg : DesignSystem::Color::kLightBg, 1.0f);
            context->Clear(bgColor);
            
            float opacity = (state->isHovered || state->isDragging) ? 0.8f : 0.3f;
            D2D1_COLOR_F color = Graphics::Engine::ColorToD2D(DesignSystem::Color::kAccent, opacity);
            
            ID2D1SolidColorBrush* brush = nullptr;
            context->CreateSolidColorBrush(color, &brush);
            
            if (brush) {
                float width = (float)(rc.right - rc.left);
                float barWidth = (state->isHovered || state->isDragging) ? width - 2.0f : 4.0f;
                float xOffset = (width - barWidth) / 2.0f;
                
                D2D1_RECT_F rect = D2D1::RectF(xOffset, thumbY, xOffset + barWidth, thumbY + thumbHeight);
                context->FillRectangle(rect, brush);
                brush->Release();
            }
            
            context->EndDraw();
        }
        
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_DESTROY:
        delete state;
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // namespace UI
