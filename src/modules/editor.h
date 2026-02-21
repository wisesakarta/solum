/*
  Saka Studio & Engineering

  Editor control functions for text manipulation, font rendering, and zoom control.
  Handles RichEdit control subclassing, word wrap, and cursor position tracking.
*/

#pragma once
#include <windows.h>
#include <string>
#include <utility>

std::wstring GetEditorText();
void SetEditorText(const std::wstring &text);
std::pair<int, int> GetCursorPos();
void ConfigureEditorControl(HWND hwnd);
void ApplyFont();
void ApplyZoom();
void ApplyWordWrap();
void ApplyEditorViewportPadding();
void DeleteWordBackward();
void DeleteWordForward();
LRESULT CALLBACK EditorSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
