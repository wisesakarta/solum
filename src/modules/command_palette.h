#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include "graphics_engine.h"

namespace UI {

struct CommandAction {
    std::wstring label;
    std::wstring description;
    UINT commandId;
    float score = 0.0f;
};

class CommandPalette {
public:
    static bool RegisterClass(HINSTANCE hInstance);
    static HWND Create(HWND hwndParent);
    static void Show(HWND hwnd);
    static void Hide(HWND hwnd);
    static bool IsVisible(HWND hwnd);
    
    static void AddCommand(const std::wstring& label, const std::wstring& description, UINT commandId);

private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static void PerformFuzzySearch(HWND hwnd, const std::wstring& query);
    
    struct State {
        HWND hwndParent;
        HWND hwndEdit = nullptr;
        HFONT hFontEdit = nullptr;
        HBRUSH hbrEdit = nullptr;
        Graphics::Engine engine;
        std::wstring query;
        std::vector<CommandAction> results;
        int selectedIndex = 0;
        float animationProgress = 0.0f; // For reveal animation
        bool isClosing = false;
    };
    
    static std::vector<CommandAction> s_commands;
};

} // namespace UI
