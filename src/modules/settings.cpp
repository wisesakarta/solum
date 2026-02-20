/*
  Saka Studio & Engineering

  Settings management for persisting user preferences via Windows Registry.
  Handles font name and font size storage and retrieval.
*/

#include "settings.h"
#include "core/globals.h"
#include "core/types.h"
#include <windows.h>
#include <algorithm>
#include <vector>

#define SETTINGS_KEY L"Software\\LegacyNotepad"
#define FONT_NAME_VALUE L"FontName"
#define FONT_SIZE_VALUE L"FontSize"
#define FONT_WEIGHT_VALUE L"FontWeight"
#define FONT_ITALIC_VALUE L"FontItalic"
#define FONT_UNDERLINE_VALUE L"FontUnderline"
#define ALWAYS_ON_TOP_VALUE L"AlwaysOnTop"
#define WORD_WRAP_VALUE L"WordWrap"
#define SHOW_STATUS_BAR_VALUE L"ShowStatusBar"
#define ZOOM_LEVEL_VALUE L"ZoomLevel"
#define THEME_VALUE L"Theme"
#define WINDOW_OPACITY_VALUE L"WindowOpacity"
#define BG_ENABLED_VALUE L"BgEnabled"
#define BG_IMAGE_PATH_VALUE L"BgImagePath"
#define BG_POSITION_VALUE L"BgPosition"
#define BG_OPACITY_VALUE L"BgOpacity"
#define CUSTOM_ICON_PATH_VALUE L"CustomIconPath"
#define CUSTOM_ICON_INDEX_VALUE L"CustomIconIndex"
#define RECENT_FILES_VALUE L"RecentFiles"
#define WINDOW_X_VALUE L"WindowX"
#define WINDOW_Y_VALUE L"WindowY"
#define WINDOW_WIDTH_VALUE L"WindowWidth"
#define WINDOW_HEIGHT_VALUE L"WindowHeight"
#define WINDOW_MAXIMIZED_VALUE L"WindowMaximized"
#define MIN_FONT_SIZE 8
#define MAX_FONT_SIZE 72
#define MIN_WINDOW_OPACITY 26
#define MIN_WINDOW_SIZE 200
#define MAX_WINDOW_SIZE 16384

static bool ReadDwordValue(HKEY hKey, const wchar_t *name, DWORD &outValue)
{
    DWORD type = 0;
    DWORD size = sizeof(outValue);
    if (RegQueryValueExW(hKey, name, nullptr, &type, reinterpret_cast<LPBYTE>(&outValue), &size) == ERROR_SUCCESS && type == REG_DWORD)
        return true;
    return false;
}

static bool ReadStringValue(HKEY hKey, const wchar_t *name, std::wstring &outValue)
{
    DWORD type = 0;
    DWORD size = 0;
    if (RegQueryValueExW(hKey, name, nullptr, &type, nullptr, &size) != ERROR_SUCCESS || type != REG_SZ || size < sizeof(wchar_t))
        return false;
    std::vector<wchar_t> buffer(size / sizeof(wchar_t));
    if (RegQueryValueExW(hKey, name, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &size) != ERROR_SUCCESS)
        return false;
    outValue.assign(buffer.data());
    return true;
}

static bool ReadIntValue(HKEY hKey, const wchar_t *name, int &outValue)
{
    DWORD dwordValue = 0;
    if (!ReadDwordValue(hKey, name, dwordValue))
        return false;
    outValue = static_cast<int>(dwordValue);
    return true;
}

static void WriteDwordValue(HKEY hKey, const wchar_t *name, DWORD value)
{
    RegSetValueExW(hKey, name, 0, REG_DWORD, reinterpret_cast<const BYTE *>(&value), sizeof(value));
}

static void WriteStringValue(HKEY hKey, const wchar_t *name, const std::wstring &value)
{
    RegSetValueExW(hKey, name, 0, REG_SZ,
                   reinterpret_cast<const BYTE *>(value.c_str()),
                   static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t)));
}

static void WriteIntValue(HKEY hKey, const wchar_t *name, int value)
{
    WriteDwordValue(hKey, name, static_cast<DWORD>(value));
}

void LoadFontSettings()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, SETTINGS_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        std::wstring strValue;
        DWORD dwordValue = 0;

        if (ReadStringValue(hKey, FONT_NAME_VALUE, strValue) && !strValue.empty())
            g_state.fontName = strValue;

        if (ReadDwordValue(hKey, FONT_SIZE_VALUE, dwordValue))
            g_state.fontSize = std::clamp(static_cast<int>(dwordValue), MIN_FONT_SIZE, MAX_FONT_SIZE);

        if (ReadDwordValue(hKey, FONT_WEIGHT_VALUE, dwordValue))
            g_state.fontWeight = static_cast<int>(dwordValue);

        if (ReadDwordValue(hKey, FONT_ITALIC_VALUE, dwordValue))
            g_state.fontItalic = (dwordValue != 0);

        if (ReadDwordValue(hKey, FONT_UNDERLINE_VALUE, dwordValue))
            g_state.fontUnderline = (dwordValue != 0);

        if (ReadDwordValue(hKey, ALWAYS_ON_TOP_VALUE, dwordValue))
            g_state.alwaysOnTop = (dwordValue != 0);

        if (ReadDwordValue(hKey, WORD_WRAP_VALUE, dwordValue))
            g_state.wordWrap = (dwordValue != 0);

        if (ReadDwordValue(hKey, SHOW_STATUS_BAR_VALUE, dwordValue))
            g_state.showStatusBar = (dwordValue != 0);

        if (ReadDwordValue(hKey, ZOOM_LEVEL_VALUE, dwordValue))
            g_state.zoomLevel = std::clamp(static_cast<int>(dwordValue), ZOOM_MIN, ZOOM_MAX);

        if (ReadDwordValue(hKey, THEME_VALUE, dwordValue))
        {
            if (dwordValue <= static_cast<DWORD>(Theme::Dark))
                g_state.theme = static_cast<Theme>(dwordValue);
        }

        if (ReadDwordValue(hKey, WINDOW_OPACITY_VALUE, dwordValue))
            g_state.windowOpacity = static_cast<BYTE>(std::clamp(static_cast<int>(dwordValue), MIN_WINDOW_OPACITY, 255));

        if (ReadDwordValue(hKey, BG_ENABLED_VALUE, dwordValue))
            g_state.background.enabled = (dwordValue != 0);

        if (ReadStringValue(hKey, BG_IMAGE_PATH_VALUE, strValue))
            g_state.background.imagePath = strValue;

        if (ReadDwordValue(hKey, BG_POSITION_VALUE, dwordValue))
        {
            if (dwordValue <= static_cast<DWORD>(BgPosition::Fill))
                g_state.background.position = static_cast<BgPosition>(dwordValue);
        }

        if (ReadDwordValue(hKey, BG_OPACITY_VALUE, dwordValue))
            g_state.background.opacity = static_cast<BYTE>(std::clamp(static_cast<int>(dwordValue), 0, 255));

        if (ReadStringValue(hKey, CUSTOM_ICON_PATH_VALUE, strValue))
            g_state.customIconPath = strValue;
        int customIconIndexValue = 0;
        if (ReadIntValue(hKey, CUSTOM_ICON_INDEX_VALUE, customIconIndexValue))
            g_state.customIconIndex = customIconIndexValue;
        else
            g_state.customIconIndex = 0;

        int intValue = 0;
        const bool hasX = ReadIntValue(hKey, WINDOW_X_VALUE, intValue);
        if (hasX)
            g_state.windowX = intValue;
        const bool hasY = ReadIntValue(hKey, WINDOW_Y_VALUE, intValue);
        if (hasY)
            g_state.windowY = intValue;

        if (ReadIntValue(hKey, WINDOW_WIDTH_VALUE, intValue))
            g_state.windowWidth = std::clamp(intValue, MIN_WINDOW_SIZE, MAX_WINDOW_SIZE);
        if (ReadIntValue(hKey, WINDOW_HEIGHT_VALUE, intValue))
            g_state.windowHeight = std::clamp(intValue, MIN_WINDOW_SIZE, MAX_WINDOW_SIZE);
        if (ReadDwordValue(hKey, WINDOW_MAXIMIZED_VALUE, dwordValue))
            g_state.windowMaximized = (dwordValue != 0);

        DWORD type = 0;
        DWORD size = 0;
        if (RegQueryValueExW(hKey, RECENT_FILES_VALUE, nullptr, &type, nullptr, &size) == ERROR_SUCCESS && type == REG_MULTI_SZ && size >= sizeof(wchar_t) * 2)
        {
            std::vector<wchar_t> buffer(size / sizeof(wchar_t));
            if (RegQueryValueExW(hKey, RECENT_FILES_VALUE, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &size) == ERROR_SUCCESS)
            {
                g_state.recentFiles.clear();
                const wchar_t *ptr = buffer.data();
                while (*ptr != L'\0')
                {
                    std::wstring value = ptr;
                    if (!value.empty())
                        g_state.recentFiles.push_back(value);
                    ptr += value.size() + 1;
                    if (g_state.recentFiles.size() >= MAX_RECENT_FILES)
                        break;
                }
            }
        }

        if (g_state.background.imagePath.empty())
            g_state.background.enabled = false;
        if (g_state.customIconPath.empty())
            g_state.customIconIndex = 0;

        RegCloseKey(hKey);
    }
}

void SaveFontSettings()
{
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, SETTINGS_KEY, 0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        WriteStringValue(hKey, FONT_NAME_VALUE, g_state.fontName);
        WriteDwordValue(hKey, FONT_SIZE_VALUE, static_cast<DWORD>(g_state.fontSize));
        WriteDwordValue(hKey, FONT_WEIGHT_VALUE, static_cast<DWORD>(g_state.fontWeight));
        WriteDwordValue(hKey, FONT_ITALIC_VALUE, g_state.fontItalic ? 1 : 0);
        WriteDwordValue(hKey, FONT_UNDERLINE_VALUE, g_state.fontUnderline ? 1 : 0);
        WriteDwordValue(hKey, ALWAYS_ON_TOP_VALUE, g_state.alwaysOnTop ? 1 : 0);
        WriteDwordValue(hKey, WORD_WRAP_VALUE, g_state.wordWrap ? 1 : 0);
        WriteDwordValue(hKey, SHOW_STATUS_BAR_VALUE, g_state.showStatusBar ? 1 : 0);
        WriteDwordValue(hKey, ZOOM_LEVEL_VALUE, static_cast<DWORD>(g_state.zoomLevel));
        WriteDwordValue(hKey, THEME_VALUE, static_cast<DWORD>(g_state.theme));
        WriteDwordValue(hKey, WINDOW_OPACITY_VALUE, static_cast<DWORD>(g_state.windowOpacity));
        WriteDwordValue(hKey, BG_ENABLED_VALUE, g_state.background.enabled ? 1 : 0);
        WriteStringValue(hKey, BG_IMAGE_PATH_VALUE, g_state.background.imagePath);
        WriteDwordValue(hKey, BG_POSITION_VALUE, static_cast<DWORD>(g_state.background.position));
        WriteDwordValue(hKey, BG_OPACITY_VALUE, static_cast<DWORD>(g_state.background.opacity));
        WriteStringValue(hKey, CUSTOM_ICON_PATH_VALUE, g_state.customIconPath);
        WriteIntValue(hKey, CUSTOM_ICON_INDEX_VALUE, g_state.customIconIndex);
        WriteIntValue(hKey, WINDOW_X_VALUE, g_state.windowX);
        WriteIntValue(hKey, WINDOW_Y_VALUE, g_state.windowY);
        WriteIntValue(hKey, WINDOW_WIDTH_VALUE, g_state.windowWidth);
        WriteIntValue(hKey, WINDOW_HEIGHT_VALUE, g_state.windowHeight);
        WriteDwordValue(hKey, WINDOW_MAXIMIZED_VALUE, g_state.windowMaximized ? 1 : 0);

        std::vector<wchar_t> multiSz;
        multiSz.reserve(2);
        for (const auto &file : g_state.recentFiles)
        {
            if (file.empty())
                continue;
            multiSz.insert(multiSz.end(), file.begin(), file.end());
            multiSz.push_back(L'\0');
            if (multiSz.size() > 65535)
                break;
        }
        if (multiSz.empty() || multiSz.back() != L'\0')
            multiSz.push_back(L'\0');
        multiSz.push_back(L'\0');
        RegSetValueExW(hKey, RECENT_FILES_VALUE, 0, REG_MULTI_SZ,
                       reinterpret_cast<const BYTE *>(multiSz.data()),
                       static_cast<DWORD>(multiSz.size() * sizeof(wchar_t)));

        RegCloseKey(hKey);
    }
}
