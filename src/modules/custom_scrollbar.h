#pragma once
#include <windows.h>
#include <d2d1_1.h>
#include "graphics_engine.h"

namespace UI {

class CustomScrollbar {
public:
    static bool RegisterClass(HINSTANCE hInstance);
    static HWND Create(HWND hwndParent, HWND hwndTarget);
    static void SetTarget(HWND hwndScrollbar, HWND hwndTarget);

private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    
    struct State {
        HWND hwndTarget;
        Graphics::Engine engine;
        bool isHovered = false;
        bool isDragging = false;
        float thumbPos = 0.0f;
        float thumbHeight = 0.0f;
        float opacity = 0.0f; // For auto-hide animation
    };
};

} // namespace UI
