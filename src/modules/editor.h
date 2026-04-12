/*
  Technical Standard

  Editor control functions for text manipulation, font rendering, and zoom control.
  Handles RichEdit control subclassing, word wrap, and cursor position tracking.
*/

#pragma once
#include <windows.h>
#include <string>
#include <utility>

std::wstring GetEditorText();
void SetEditorText(const std::wstring &text);
std::string GetEditorRichText();
void SetEditorRichText(const std::string &rtf);
std::pair<int, int> GetCursorPos();
DWORD BuildEditorWindowStyle();
void ConfigureEditorControl(HWND hwnd);
void ApplyFont();
void ApplyZoom();
bool ScrollEditorFromMouseWheel(HWND hwndEditor, WPARAM wParam);
void ApplyWordWrap();
void ApplyEditorViewportPadding();
void ApplyEditorScrollbarChrome();
void DeleteWordBackward();
void DeleteWordForward();
LRESULT CALLBACK EditorSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
