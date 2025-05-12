// ==WindhawkMod==
// @id              system-metrics-loader
// @name            System Metrics Loader
// @description     Loads system metrics from active theme or NONCLIENTMETRICSA on theme change, mimicking Windows 7
// @version         2.7
// @author          MrGRiM
// @github          https://github.com/MrGRiM01
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lole32 -luxtheme -lgdi32 -luser32 -ldwmapi
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# System Metrics Loader
Loads system metrics from the active msstyles theme or NONCLIENTMETRICSA on theme change, restoring Windows 7 behavior.
No theme modifications required.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- loadThemeMetrics: true
  $name: Load Theme Metrics
  $description: Enable loading system metrics from active theme or NONCLIENTMETRICSA
- useThemeMetrics: true
  $name: Prefer Theme Metrics
  $description: Prioritize theme metrics; falls back to NONCLIENTMETRICSA if unavailable
- allowSystemMetrics: true
  $name: Allow NONCLIENTMETRICSA
  $description: Use NONCLIENTMETRICSA as fallback for missing theme metrics
*/
// ==/WindhawkModSettings==

#include <windhawk_api.h>
#include <windhawk_utils.h>
#include <windows.h>
#include <uxtheme.h>
#include <vssym32.h>

struct Settings {
    bool loadThemeMetrics;
    bool useThemeMetrics;
    bool allowSystemMetrics;
} settings;

static HANDLE g_themeChangeEvent = nullptr;
static HANDLE g_themeChangeThread = nullptr;
static DWORD lastActionTime = 0;
static WCHAR lastThemeName[MAX_PATH] = {0};
static const DWORD ACTION_THROTTLE_MS = 5000;
static const DWORD THEME_CHANGE_DELAY_MS = 1000;

typedef void (WINAPI *SetThemeAppProperties_t)(DWORD);

// Convert pixels to twips (1 pixel = -15 twips at 96 DPI)
int PixelsToTwips(int pixels) {
    return pixels * -15;
}

bool EnsureRegistryKey(const wchar_t* subKey, REGSAM access, HKEY* hKey) {
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, subKey, 0, access | KEY_WOW64_64KEY, hKey);
    if (result == ERROR_SUCCESS) return true;

    result = RegCreateKeyExW(HKEY_CURRENT_USER, subKey, 0, nullptr, REG_OPTION_NON_VOLATILE, access | KEY_WOW64_64KEY, nullptr, hKey, nullptr);
    if (result == ERROR_SUCCESS) {
        Wh_Log(L"Created registry key: %s", subKey);
        return true;
    }
    Wh_Log(L"Failed to create registry key: %s, error=%d", subKey, result);
    return false;
}

bool LoadMetricsFromMsstyles(int* metrics) {
    WCHAR themeName[MAX_PATH];
    HRESULT hr = GetCurrentThemeName(themeName, MAX_PATH, nullptr, 0, nullptr, 0);
    if (FAILED(hr)) {
        Wh_Log(L"Failed to get theme name, error=0x%X", hr);
        return false;
    }
    Wh_Log(L"Active theme: %s", themeName);

    HTHEME hTheme = OpenThemeData(nullptr, L"WINDOW");
    bool themeOpened = !!hTheme;
    if (!themeOpened) {
        Wh_Log(L"Failed to open theme for WINDOW");
    }

    bool anySuccess = false;
    int sysSizeIds[] = { SM_CYSIZE, SM_CXFRAME, SM_CXPADDEDBORDER, SM_CYMENUSIZE, SM_CXVSCROLL, SM_CXVSCROLL, SM_CYSMSIZE, SM_CXSIZE, SM_CXSMSIZE, SM_CXMENUSIZE };
    const wchar_t* metricNames[] = { L"CaptionHeight", L"BorderWidth", L"PaddedBorderWidth", L"MenuHeight", L"ScrollWidth", L"ScrollHeight", L"SmCaptionHeight", L"CaptionWidth", L"SmCaptionWidth", L"MenuWidth" };

    if (themeOpened && settings.useThemeMetrics) {
        for (int i = 0; i < 10; i++) {
            metrics[i] = GetThemeSysSize(hTheme, sysSizeIds[i]);
            if (metrics[i] != 0) {
                metrics[i] = PixelsToTwips(metrics[i]);
                Wh_Log(L"GetThemeSysSize: %s (SM_%d)=%d (twips=%d)", metricNames[i], sysSizeIds[i], metrics[i] / -15, metrics[i]);
                anySuccess = true;
            } else {
                Wh_Log(L"GetThemeSysSize failed: %s (SM_%d)=0", metricNames[i], sysSizeIds[i]);
            }
        }
        CloseThemeData(hTheme);
    }

    if (settings.allowSystemMetrics) {
        NONCLIENTMETRICSA ncm = { sizeof(NONCLIENTMETRICSA) };
        if (SystemParametersInfoA(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0)) {
            metrics[0] = metrics[0] == 0 ? PixelsToTwips(ncm.iCaptionHeight) : metrics[0];
            metrics[1] = metrics[1] == 0 ? PixelsToTwips(ncm.iBorderWidth) : metrics[1];
            metrics[2] = metrics[2] == 0 ? PixelsToTwips(ncm.iPaddedBorderWidth) : metrics[2];
            metrics[3] = metrics[3] == 0 ? PixelsToTwips(ncm.iMenuHeight) : metrics[3];
            metrics[4] = metrics[4] == 0 ? PixelsToTwips(ncm.iScrollWidth) : metrics[4];
            metrics[5] = metrics[5] == 0 ? PixelsToTwips(ncm.iScrollHeight) : metrics[5];
            metrics[6] = metrics[6] == 0 ? PixelsToTwips(ncm.iSmCaptionHeight) : metrics[6];
            metrics[7] = metrics[7] == 0 ? PixelsToTwips(ncm.iCaptionWidth) : metrics[7];
            metrics[8] = metrics[8] == 0 ? PixelsToTwips(ncm.iSmCaptionWidth) : metrics[8];
            metrics[9] = metrics[9] == 0 ? PixelsToTwips(ncm.iMenuWidth) : metrics[9];
            Wh_Log(L"NONCLIENTMETRICSA: CaptionHeight=%d, BorderWidth=%d, PaddedBorder=%d, MenuHeight=%d, ScrollWidth=%d, ScrollHeight=%d, SmCaptionHeight=%d, CaptionWidth=%d, SmCaptionWidth=%d, MenuWidth=%d",
                   ncm.iCaptionHeight, ncm.iBorderWidth, ncm.iPaddedBorderWidth, ncm.iMenuHeight, ncm.iScrollWidth, ncm.iScrollHeight, ncm.iSmCaptionHeight, ncm.iCaptionWidth, ncm.iSmCaptionWidth, ncm.iMenuWidth);
            anySuccess = true;
        } else {
            Wh_Log(L"Failed to get NONCLIENTMETRICSA, error=%d", GetLastError());
        }
    }

    if (anySuccess) {
        Wh_Log(L"Loaded metrics: CaptionHeight=%d, BorderWidth=%d, PaddedBorder=%d, MenuHeight=%d, ScrollWidth=%d, ScrollHeight=%d, SmCaptionHeight=%d, CaptionWidth=%d, SmCaptionWidth=%d, MenuWidth=%d",
               metrics[0], metrics[1], metrics[2], metrics[3], metrics[4], metrics[5], metrics[6], metrics[7], metrics[8], metrics[9]);
    } else {
        Wh_Log(L"Failed to load metrics from %s or NONCLIENTMETRICSA", themeName);
    }
    return anySuccess;
}

bool ApplyMetrics(const int* metrics) {
    HKEY hKey;
    const wchar_t* subKey = L"Control Panel\\Desktop\\WindowMetrics";
    if (!EnsureRegistryKey(subKey, KEY_SET_VALUE, &hKey)) {
        Wh_Log(L"Failed to ensure registry key: %s", subKey);
        return false;
    }

    struct MetricEntry {
        const wchar_t* name;
        int value;
    };

    MetricEntry entries[] = {
        { L"CaptionHeight", metrics[0] },
        { L"CaptionWidth", metrics[7] },
        { L"BorderWidth", metrics[1] },
        { L"PaddedBorderWidth", metrics[2] },
        { L"MenuHeight", metrics[3] },
        { L"MenuWidth", metrics[9] },
        { L"ScrollWidth", metrics[4] },
        { L"ScrollHeight", metrics[5] },
        { L"SmCaptionHeight", metrics[6] },
        { L"SmCaptionWidth", metrics[8] }
    };

    bool success = false;
    for (const auto& entry : entries) {
        if (entry.value == 0) {
            Wh_Log(L"Skipping %s, value=0", entry.name);
            continue;
        }
        WCHAR value[32];
        _itow_s(entry.value, value, 10);
        LONG result = RegSetValueExW(hKey, entry.name, 0, REG_SZ, (BYTE*)value, (wcslen(value) + 1) * sizeof(WCHAR));
        if (result == ERROR_SUCCESS) {
            Wh_Log(L"Set %s=%s", entry.name, value);
            success = true;
        } else {
            Wh_Log(L"Failed to set %s=%s, error=%d", entry.name, value, result);
        }
    }
    RegCloseKey(hKey);

    NONCLIENTMETRICSW ncm = { sizeof(NONCLIENTMETRICSW) };
    if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0)) {
        if (metrics[0] != 0) ncm.iCaptionHeight = metrics[0] / -15;
        if (metrics[7] != 0) ncm.iCaptionWidth = metrics[7] / -15;
        if (metrics[1] != 0) ncm.iBorderWidth = metrics[1] / -15;
        if (metrics[2] != 0) ncm.iPaddedBorderWidth = metrics[2] / -15;
        if (metrics[3] != 0) ncm.iMenuHeight = metrics[3] / -15;
        if (metrics[9] != 0) ncm.iMenuWidth = metrics[9] / -15;
        if (metrics[4] != 0) ncm.iScrollWidth = metrics[4] / -15;
        if (metrics[5] != 0) ncm.iScrollHeight = metrics[5] / -15;
        if (metrics[6] != 0) ncm.iSmCaptionHeight = metrics[6] / -15;
        if (metrics[8] != 0) ncm.iSmCaptionWidth = metrics[8] / -15;

        if (!SystemParametersInfoW(SPI_SETNONCLIENTMETRICS, sizeof(ncm), &ncm, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)) {
            Wh_Log(L"Failed to set NONCLIENTMETRICS, error=%d", GetLastError());
        } else {
            Wh_Log(L"Applied NONCLIENTMETRICS");
            success = true;
        }
    } else {
        Wh_Log(L"Failed to get NONCLIENTMETRICS, error=%d", GetLastError());
    }
    return success;
}

void NotifyWindows() {
    SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0, (LPARAM)L"WindowMetrics", SMTO_ABORTIFHUNG, 2000, nullptr);
    SendMessageTimeoutW(HWND_BROADCAST, WM_WININICHANGE, 0, (LPARAM)L"WindowMetrics", SMTO_ABORTIFHUNG, 2000, nullptr);
    SendMessageTimeoutW(HWND_BROADCAST, WM_THEMECHANGED, 0, 0, SMTO_ABORTIFHUNG, 2000, nullptr);
    Wh_Log(L"Broadcast metric change notifications");
}

bool ApplyThemeMetrics() {
    DWORD currentTime = GetTickCount();
    if (currentTime - lastActionTime < ACTION_THROTTLE_MS) {
        Wh_Log(L"Skipping action, last executed %u ms ago", currentTime - lastActionTime);
        return false;
    }

    WCHAR themeName[MAX_PATH];
    HRESULT hr = GetCurrentThemeName(themeName, MAX_PATH, nullptr, 0, nullptr, 0);
    if (FAILED(hr)) {
        Wh_Log(L"Failed to get theme name, error=0x%X", hr);
        return false;
    }

    if (wcscmp(themeName, lastThemeName) == 0) {
        Wh_Log(L"Theme unchanged: %s, skipping", themeName);
        return false;
    }

    int metrics[10] = { 0 };
    if (settings.loadThemeMetrics && LoadMetricsFromMsstyles(metrics)) {
        if (ApplyMetrics(metrics)) {
            wcscpy_s(lastThemeName, themeName);
            NotifyWindows();
            lastActionTime = currentTime;
            Wh_Log(L"Applied metrics for theme: %s", themeName);
            return true;
        } else {
            Wh_Log(L"Failed to apply metrics");
        }
    } else {
        Wh_Log(L"No metrics applied");
    }
    return false;
}

void (*originalSetThemeAppProperties)(DWORD);
void WINAPI HookedSetThemeAppProperties(DWORD dwFlags) {
    Wh_Log(L"SetThemeAppProperties called, dwFlags=0x%X", dwFlags);
    originalSetThemeAppProperties(dwFlags);
    if (settings.loadThemeMetrics && (dwFlags & STAP_ALLOW_CONTROLS)) {
        Sleep(THEME_CHANGE_DELAY_MS);
        ApplyThemeMetrics();
    }
}

DWORD WINAPI ThemeChangeWatcher(LPVOID) {
    HKEY hKey;
    const wchar_t* themeKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
    if (RegOpenKeyExW(HKEY_CURRENT_USER, themeKey, 0, KEY_NOTIFY, &hKey) != ERROR_SUCCESS) {
        Wh_Log(L"Failed to open Themes\\Personalize key, error=%d", GetLastError());
        return 1;
    }

    while (true) {
        RegNotifyChangeKeyValue(hKey, TRUE, REG_NOTIFY_CHANGE_LAST_SET, g_themeChangeEvent, TRUE);
        if (WaitForSingleObject(g_themeChangeEvent, INFINITE) == WAIT_OBJECT_0) {
            Wh_Log(L"Theme change detected");
            if (settings.loadThemeMetrics) {
                Sleep(THEME_CHANGE_DELAY_MS);
                ApplyThemeMetrics();
            }
            ResetEvent(g_themeChangeEvent);
        } else {
            Wh_Log(L"Theme watcher wait failed, error=%d", GetLastError());
            break;
        }
    }
    RegCloseKey(hKey);
    return 0;
}

void StartThemeChangeWatcher() {
    if (g_themeChangeEvent || g_themeChangeThread) return;
    g_themeChangeEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!g_themeChangeEvent) {
        Wh_Log(L"Failed to create theme change event, error=%d", GetLastError());
        return;
    }
    g_themeChangeThread = CreateThread(nullptr, 0, ThemeChangeWatcher, nullptr, 0, nullptr);
    if (!g_themeChangeThread) {
        Wh_Log(L"Failed to create theme change thread, error=%d", GetLastError());
        CloseHandle(g_themeChangeEvent);
        g_themeChangeEvent = nullptr;
    } else {
        Wh_Log(L"Started theme change watcher");
    }
}

void StopThemeChangeWatcher() {
    if (g_themeChangeThread) {
        TerminateThread(g_themeChangeThread, 0);
        CloseHandle(g_themeChangeThread);
        g_themeChangeThread = nullptr;
    }
    if (g_themeChangeEvent) {
        CloseHandle(g_themeChangeEvent);
        g_themeChangeEvent = nullptr;
    }
    Wh_Log(L"Stopped theme change watcher");
}

void LoadSettings() {
    settings.loadThemeMetrics = Wh_GetIntSetting(L"loadThemeMetrics");
    settings.useThemeMetrics = Wh_GetIntSetting(L"useThemeMetrics");
    settings.allowSystemMetrics = Wh_GetIntSetting(L"allowSystemMetrics");
    Wh_Log(L"Loaded settings: loadThemeMetrics=%d, useThemeMetrics=%d, allowSystemMetrics=%d",
           settings.loadThemeMetrics, settings.useThemeMetrics, settings.allowSystemMetrics);
}

void Wh_ModSettingsChanged() {
    LoadSettings();
    if (settings.loadThemeMetrics) ApplyThemeMetrics();
}

BOOL Wh_ModInit() {
    Wh_Log(L"Initializing System Metrics Loader");
    LoadSettings();

    WCHAR processName[MAX_PATH] = {0};
    GetModuleFileNameW(nullptr, processName, MAX_PATH);
    if (_wcsicmp(wcsrchr(processName, L'\\') + 1, L"explorer.exe") != 0) {
        Wh_Log(L"Non-Explorer process, applying metrics if enabled");
        if (settings.loadThemeMetrics && settings.useThemeMetrics) ApplyThemeMetrics();
        return TRUE;
    }

    HMODULE hUxtheme = LoadLibraryW(L"uxtheme.dll");
    if (!hUxtheme) {
        Wh_Log(L"Failed to load uxtheme.dll");
        return FALSE;
    }

    if (settings.loadThemeMetrics && settings.useThemeMetrics) {
        if (!WindhawkUtils::SetFunctionHook(
            (SetThemeAppProperties_t)GetProcAddress(hUxtheme, "SetThemeAppProperties"),
            HookedSetThemeAppProperties,
            &originalSetThemeAppProperties)) {
            Wh_Log(L"Failed to hook SetThemeAppProperties");
            return FALSE;
        }
        Wh_Log(L"Hooked SetThemeAppProperties");
        StartThemeChangeWatcher();
        ApplyThemeMetrics();
    }
    return TRUE;
}

void Wh_ModUninit() {
    Wh_Log(L"Uninitializing System Metrics Loader");
    StopThemeChangeWatcher();
    if (originalSetThemeAppProperties) {
        if (Wh_RemoveFunctionHook((void*)originalSetThemeAppProperties)) {
            Wh_Log(L"Removed SetThemeAppProperties hook");
        } else {
            Wh_Log(L"Failed to remove SetThemeAppProperties hook (may be restricted during uninit)");
        }
    }
}