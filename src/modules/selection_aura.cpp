#include "selection_aura.h"
#include "core/globals.h"
#include "design_system.h"
#include "editor.h"
#include <richedit.h>
#include <algorithm>

namespace UI {

bool SelectionAura::RegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"TechnicalStandardSelectionAura";
    wc.hCursor = LoadCursor(nullptr, IDC_IBEAM);
    // WS_EX_TRANSPARENT + hbrBackground = NULL allows click-through
    wc.hbrBackground = nullptr;
    return ::RegisterClassExW(&wc) != 0;
}

HWND SelectionAura::Create(HWND hwndParent, HWND hwndEditor)
{
    HWND hwnd = CreateWindowExW(WS_EX_TRANSPARENT | WS_EX_LAYERED, L"TechnicalStandardSelectionAura", nullptr,
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0,
        hwndParent, nullptr, GetModuleHandle(nullptr), nullptr);
    
    if (hwnd) {
        SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA);
        State* state = new State();
        state->hwndEditor = hwndEditor;
        state->engine.Initialize(hwnd);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    }
    return hwnd;
}

void SelectionAura::Update(HWND hwnd)
{
    State* state = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) return;

    CHARRANGE cr;
    SendMessageW(state->hwndEditor, EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&cr));
    
    state->selectionRectCount = 0;
    if (cr.cpMin != cr.cpMax) {
        // Simple bounding box logic for now (Renaissance Step 1)
        POINTL ptMin, ptMax;
        SendMessageW(state->hwndEditor, EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&ptMin), cr.cpMin);
        SendMessageW(state->hwndEditor, EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&ptMax), cr.cpMax);
        
        // Convert editor client to aura client (which should be same)
        state->selectionRects[0] = { (LONG)ptMin.x, (LONG)ptMin.y, (LONG)ptMax.x, (LONG)ptMax.y + 20 };
        state->selectionRectCount = 1;
    }
    
    InvalidateRect(hwnd, nullptr, FALSE);
}

void SelectionAura::SetEditor(HWND hwndAura, HWND hwndEditor)
{
    if (!hwndAura)
        return;
    State* state = reinterpret_cast<State*>(GetWindowLongPtrW(hwndAura, GWLP_USERDATA));
    if (!state)
        return;
    state->hwndEditor = hwndEditor;
    state->selectionRectCount = 0;
    InvalidateRect(hwndAura, nullptr, FALSE);
}

LRESULT CALLBACK SelectionAura::WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    State* state = reinterpret_cast<State*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg) {
    case WM_SIZE:
        if (state) state->engine.Resize(LOWORD(lParam), HIWORD(lParam));
        return 0;
    case WM_PAINT: {
        if (!state) return 0;
        PAINTSTRUCT ps;
        BeginPaint(hwnd, &ps);
        
        auto context = state->engine.GetDeviceContext();
        if (context) {
            context->BeginDraw();
            context->Clear(nullptr); // Fully transparent bg
            
            if (state->selectionRectCount > 0) {
                D2D1_COLOR_F glowColor = Graphics::Engine::ColorToD2D(DesignSystem::Color::kAccent, 0.4f);
                
                for (int i = 0; i < state->selectionRectCount; ++i) {
                    RECT& r = state->selectionRects[i];
                    D2D1_RECT_F rect = D2D1::RectF((float)r.left, (float)r.top, (float)r.right, (float)r.bottom);
                    state->engine.DrawBlurredRect(rect, glowColor, 4.0f);
                }
            }
            context->EndDraw();
        }
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_NCHITTEST: return HTTRANSPARENT; // Ensure mouse passes through to editor
    case WM_MOUSEWHEEL:
        if (state && state->hwndEditor)
        {
            ScrollEditorFromMouseWheel(state->hwndEditor, wParam);
            if (g_hwndScrollbar)
                InvalidateRect(g_hwndScrollbar, nullptr, FALSE);
            return 0;
        }
        return 0;
    case WM_DESTROY:
        delete state;
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // namespace UI
