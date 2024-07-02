#include <windows.h>
#include <oleacc.h>
#include <psapi.h>
#include <iostream>
#include <string>

#pragma comment(lib, "Oleacc.lib")
#pragma comment(lib, "Psapi.lib")

// Global hook handle
HHOOK hMouseHook = NULL;

// Function prototypes
std::string ConvertWStringToString(const std::wstring& wstr);
std::string GetProcessName(HWND hwnd);
std::string GetAccessibleRole(IAccessible* pAcc, VARIANT varChild);
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam);
void SetMouseHook();
void UnhookMouse();

// Helper function to convert wstring to string
std::string ConvertWStringToString(const std::wstring& wstr) {
    int bufferSize = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string str(bufferSize, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], bufferSize, nullptr, nullptr);
    return str;
}

// Helper function to get the process name by HWND
std::string GetProcessName(HWND hwnd) {
    DWORD processId;
    GetWindowThreadProcessId(hwnd, &processId);
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
    if (hProcess) {
        TCHAR exeName[MAX_PATH];
        if (GetModuleFileNameEx(hProcess, NULL, exeName, MAX_PATH)) {
            CloseHandle(hProcess);
            std::wstring exeWStr = exeName;
            std::string exeStr = ConvertWStringToString(exeWStr);
            size_t pos = exeStr.find_last_of("\\/");
            return (pos != std::string::npos) ? exeStr.substr(pos + 1) : exeStr;
        }
        CloseHandle(hProcess);
    }
    return "unknown.exe";
}

// Helper function to get the role of the accessible object
std::string GetAccessibleRole(IAccessible* pAcc, VARIANT varChild) {
    VARIANT varRole;
    VariantInit(&varRole);
    if (pAcc!=NULL && pAcc->get_accRole(varChild, &varRole) == S_OK) {
        if (varRole.vt == VT_I4) {
            switch (varRole.lVal) {
            case 8: return "alert";            // ROLE_SYSTEM_ALERT
            case 43: return "button";          // ROLE_SYSTEM_BUTTON
            case 44: return "check box";       // ROLE_SYSTEM_CHECKBUTTON
            case 46: return "combo box";       // ROLE_SYSTEM_COMBOBOX
            case 15: return "document";        // ROLE_SYSTEM_DOCUMENT
            case 30: return "hyperlink";       // ROLE_SYSTEM_LINK
            case 34: return "list";            // ROLE_SYSTEM_LIST
            case 35: return "list item";       // ROLE_SYSTEM_LISTITEM
            case 45: return "menu item";       // ROLE_SYSTEM_MENUITEM
            case 48: return "progress bar";    // ROLE_SYSTEM_PROGRESSBAR
            case 50: return "radio button";    // ROLE_SYSTEM_RADIOBUTTON
            case 53: return "scroll bar";      // ROLE_SYSTEM_SCROLLBAR
            case 21: return "separator";       // ROLE_SYSTEM_SEPARATOR
            case 52: return "slider";          // ROLE_SYSTEM_SLIDER
            case 54: return "spin box";        // ROLE_SYSTEM_SPINBUTTON
            case 41: return "static text";     // ROLE_SYSTEM_STATICTEXT
            case 37: return "tab";             // ROLE_SYSTEM_TAB
            case 38: return "table";           // ROLE_SYSTEM_TABLE
            case 42: return "text box";        // ROLE_SYSTEM_TEXT
            case 1: return "title bar";        // ROLE_SYSTEM_TITLEBAR
            case 10: return "client";          // ROLE_SYSTEM_CLIENT
            default: return "unknown role";
            }
        }
    }
    VariantClear(&varRole);
    return "unknown role";
}

// Function to get the IAccessible object at the cursor position
IAccessible* GetAccessibleObjectFromPoint(POINT pt, VARIANT* pVarChild) {
    IAccessible* pAcc = nullptr;
    HRESULT hr = AccessibleObjectFromPoint(pt, &pAcc, pVarChild);
    if (SUCCEEDED(hr)) {
        return pAcc;
    }
    return nullptr;
}

// Hook procedure (callback function)
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        MSLLHOOKSTRUCT* pMouseStruct = (MSLLHOOKSTRUCT*)lParam;
        if (pMouseStruct != nullptr) {
            if (wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN) {
                POINT pt = pMouseStruct->pt;
                HWND hwnd = WindowFromPoint(pt);
                std::string processName = GetProcessName(hwnd);

                VARIANT varChild;
                IAccessible* pAcc = GetAccessibleObjectFromPoint(pt, &varChild);
                std::string elementType = GetAccessibleRole(pAcc, varChild);
                if (pAcc) {
                    pAcc->Release();  // Release the IAccessible object
                }

                std::cout << processName << ": {X=" << pt.x << ", Y=" << pt.y << "}: " << elementType << std::endl;
            }
        }
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

// Function to set the mouse hook
void SetMouseHook() {
    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, LowLevelMouseProc, GetModuleHandle(NULL), 0);
    if (hMouseHook == NULL) {
        std::cerr << "Failed to install mouse hook!" << std::endl;
    }
}

// Function to unhook the mouse
void UnhookMouse() {
    UnhookWindowsHookEx(hMouseHook);
}

int main() {
    HRESULT hr = CoInitialize(NULL);  // Initialize COM library
    if (FAILED(hr)) {
        std::cerr << "Failed to initialize COM library!" << std::endl;
        return 1;
    }

    SetMouseHook();
    MSG msg;
    // Message loop to keep the application running
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    UnhookMouse();
    CoUninitialize();  // Uninitialize COM library
    return 0;
}
