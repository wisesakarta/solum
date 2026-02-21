/*
  Saka Studio & Engineering

  Settings management for persisting user preferences via Windows Registry.
  Handles font name and font size storage and retrieval.
*/

#pragma once

#include <string>
#include <vector>

void LoadFontSettings();
void SaveFontSettings();
void SaveOpenTabsSession(const std::vector<std::wstring> &tabPaths, int activeTabIndex);
void LoadOpenTabsSession(std::vector<std::wstring> &tabPaths, int &activeTabIndex);
