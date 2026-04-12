#pragma once
#include <windows.h>
#include "graphics_engine.h"

namespace UI {

class SelectionAura {
public:
    static bool RegisterClass(HINSTANCE hInstance);
    static HWND Create(HWND hwndParent, HWND hwndEditor);
    static void Update(HWND hwnd);
    static void SetEditor(HWND hwndAura, HWND hwndEditor);

private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    
    struct State {
        HWND hwndEditor;
        Graphics::Engine engine;
        RECT selectionRects[16];
        int selectionRectCount = 0;
    };
};

} // namespace UI
