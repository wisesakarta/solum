/*
  Saka Studio & Engineering

  Theme management module providing dark mode support and visual style functions.
  Handles Windows immersive dark mode APIs and applies theme to all UI elements.
*/

#pragma once

#include <windows.h>

bool IsDarkMode();
bool SetTitleBarDark(HWND hwnd, BOOL dark);
void ApplyThemeToWindowTree(HWND hwnd);
void ApplyTheme();
void ToggleDarkMode();
LRESULT CALLBACK StatusSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

COLORREF ThemeColorEditorBackground(bool dark);
COLORREF ThemeColorEditorText(bool dark);
COLORREF ThemeColorStatusBackground(bool dark);
COLORREF ThemeColorStatusText(bool dark);
COLORREF ThemeColorMenuBackground(bool dark);
COLORREF ThemeColorMenuHoverBackground(bool dark);
COLORREF ThemeColorMenuText(bool dark);
COLORREF ThemeColorChromeBorder(bool dark);
