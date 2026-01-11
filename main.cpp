#include <windows.h>
#include <shellapi.h>
#include <taskschd.h>
#include <comdef.h>
#include "resource.h"

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsuppw.lib")

// 配置区域 ====================================================
// 在这里修改配置即可改变程序行为
struct AppConfig {
    // 侧键配置
    BYTE sideButton = XBUTTON2;  // 侧键选择：XBUTTON1 或 XBUTTON2

    // 侧键单独按下时发送的键
    BYTE sideButtonAlone = VK_MEDIA_PLAY_PAUSE;  // 播放/暂停

    // 侧键+左键的功能
    BYTE sideWithLeft = VK_MEDIA_PREV_TRACK;     // 上一曲

    // 侧键+右键的功能  
    BYTE sideWithRight = VK_MEDIA_NEXT_TRACK;    // 下一曲

    // 侧键+滚轮向上的功能
    BYTE sideWithWheelUp = VK_VOLUME_UP;         // 音量加

    // 侧键+滚轮向下的功能
    BYTE sideWithWheelDown = VK_VOLUME_DOWN;     // 音量减

    // 是否显示托盘图标
	bool showTrayIcon = true;

    // 托盘图标提示文本
    wchar_t trayTooltip[128] = L"MouseTune - 媒体控制工具";

    // 以下是可用的媒体键代码（供参考）：
    // VK_MEDIA_PLAY_PAUSE   (0xB3) - 播放/暂停
    // VK_MEDIA_NEXT_TRACK   (0xB0) - 下一曲
    // VK_MEDIA_PREV_TRACK   (0xB1) - 上一曲
    // VK_MEDIA_STOP         (0xB2) - 停止
    // VK_VOLUME_UP          (0xAF) - 音量加
    // VK_VOLUME_DOWN        (0xAE) - 音量减
    // VK_VOLUME_MUTE        (0xAD) - 静音
};
// =============================================================

// 全局配置对象
AppConfig g_config;

// 系统常量定义
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_AUTOSTART 1002
#define ID_TRAY_TOGGLE 1003

// 计划任务名称
const wchar_t* TASK_NAME = L"MouseTune_AutoStart";

// 1定义全局变量存储消息 ID
UINT g_uMsgTaskbarCreated = 0;

// 定义全局变量存储启用状态
bool g_isEnabled = true;

bool IsRunAsAdmin() {
    BOOL fIsRunAsAdmin = FALSE;
    PSID AdministrationGroup;
    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
    if (AllocateAndInitializeSid(&NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID,
        DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &AdministrationGroup)) {
        CheckTokenMembership(NULL, AdministrationGroup, &fIsRunAsAdmin);
        FreeSid(AdministrationGroup);
    }
    return fIsRunAsAdmin == TRUE;
}
void ElevateNow() {
    wchar_t szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, MAX_PATH);
    SHELLEXECUTEINFO sei = { sizeof(sei) };
    sei.lpVerb = L"runas";
    sei.lpFile = szPath;
    sei.lpParameters = L"--autostart";
    sei.hwnd = NULL;
    sei.nShow = SW_NORMAL;
    if (ShellExecuteEx(&sei)) {
        PostQuitMessage(0); // 启动新进程后关闭旧进程
    }
}
// 计划任务管理 (创建或删除)
bool ManageStartupTask(bool enable) {
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) return false;

    bool success = false;
    ITaskService* pService = NULL;
    ITaskFolder* pRootFolder = NULL;

    do {
        // 1. 连接到计划任务服务
        if (FAILED(CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService))) break;
        if (FAILED(pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()))) break;
        if (FAILED(pService->GetFolder(_bstr_t(L"\\"), &pRootFolder))) break;

        if (!enable) {
            // 删除任务
            pRootFolder->DeleteTask(_bstr_t(TASK_NAME), 0);
            success = true;
            break;
        }

        // 2. 创建任务定义
        ITaskDefinition* pTask = NULL;
        if (FAILED(pService->NewTask(0, &pTask))) break;

        // 设置登录时启动
        ITriggerCollection* pTriggers = NULL;
        pTask->get_Triggers(&pTriggers);
        ITrigger* pTrigger = NULL;
        pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);
        pTriggers->Release();
        pTrigger->Release();

        // 设置动作：执行程序
        IActionCollection* pActions = NULL;
        pTask->get_Actions(&pActions);
        IAction* pAction = NULL;
        pActions->Create(TASK_ACTION_EXEC, &pAction);
        IExecAction* pExecAction = NULL;
        pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);

        wchar_t szPath[MAX_PATH];
        GetModuleFileNameW(NULL, szPath, MAX_PATH);
        pExecAction->put_Path(_bstr_t(szPath));
        pExecAction->Release();
        pAction->Release();
        pActions->Release();

        // 关键：设置管理员权限运行 (HighestAvailable)
        IPrincipal* pPrincipal = NULL;
        pTask->get_Principal(&pPrincipal);
        pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
        pPrincipal->Release();

        // 设置登录后延迟 5-10 秒启动
        ILogonTrigger* pLogonTrigger = NULL;
        pTrigger->QueryInterface(IID_ILogonTrigger, (void**)&pLogonTrigger);
        pLogonTrigger->put_Delay(_bstr_t(L"PT5S")); // PT5S 表示延迟 5 秒 (ISO 8601 格式)
        pLogonTrigger->Release();

        // 注册任务
        IRegisteredTask* pRegisteredTask = NULL;
        hr = pRootFolder->RegisterTaskDefinition(_bstr_t(TASK_NAME), pTask, TASK_CREATE_OR_UPDATE,
            _variant_t(), _variant_t(), TASK_LOGON_INTERACTIVE_TOKEN, _variant_t(L""), &pRegisteredTask);

        if (SUCCEEDED(hr)) success = true;

        pTask->Release();
        if (pRegisteredTask) pRegisteredTask->Release();

    } while (0);

    if (pRootFolder) pRootFolder->Release();
    if (pService) pService->Release();
    CoUninitialize();
    return success;
}
// 检查任务是否存在（用于菜单打钩）
bool IsTaskRegistered() {
    bool found = false;
    ITaskService* pService = NULL;
    ITaskFolder* pRootFolder = NULL;
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (SUCCEEDED(CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService))) {
        if (SUCCEEDED(pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()))) {
            if (SUCCEEDED(pService->GetFolder(_bstr_t(L"\\"), &pRootFolder))) {
                IRegisteredTask* pTask = NULL;
                if (SUCCEEDED(pRootFolder->GetTask(_bstr_t(TASK_NAME), &pTask))) {
                    found = true;
                    pTask->Release();
                }
            }
        }
    }
    if (pRootFolder) pRootFolder->Release();
    if (pService) pService->Release();
    CoUninitialize();
    return found;
}

// 全局变量
HHOOK hMouseHook = NULL;
bool isSideButtonPressed = false;
bool isInteracted = false;
NOTIFYICONDATA nid = { 0 };

// 发送按键函数
void SendKey(BYTE vkey) {
    INPUT input = { 0 };
    input.type = INPUT_KEYBOARD;
    input.ki.wVk = vkey;
    SendInput(1, &input, sizeof(INPUT));

    input.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &input, sizeof(INPUT));
}

// 鼠标钩子回调函数
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (!g_isEnabled) {
        return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
    }
    if (nCode >= 0) {
        MSLLHOOKSTRUCT* pMouse = (MSLLHOOKSTRUCT*)lParam;

        // 1. 检测侧键按下
        if (wParam == WM_XBUTTONDOWN) {
            if (HIWORD(pMouse->mouseData) == g_config.sideButton) {
                isSideButtonPressed = true;
                isInteracted = false;
                return 1;
            }
        }
        // 2. 检测侧键释放
        else if (wParam == WM_XBUTTONUP) {
            if (HIWORD(pMouse->mouseData) == g_config.sideButton) {
                isSideButtonPressed = false;
                // 如果在侧键按下期间没有交互，则触发单独侧键功能
                if (!isInteracted) {
                    SendKey(g_config.sideButtonAlone);
                }
                return 1;
            }
        }

        // 3. 侧键按下时的组合功能
        if (isSideButtonPressed) {
            // 侧键 + 左键
            if (wParam == WM_LBUTTONDOWN) {
                isInteracted = true;
                SendKey(g_config.sideWithLeft);
                return 1;
            }
            if (wParam == WM_LBUTTONUP) return 1;

            // 侧键 + 右键
            if (wParam == WM_RBUTTONDOWN) {
                isInteracted = true;
                SendKey(g_config.sideWithRight);
                return 1;
            }
            if (wParam == WM_RBUTTONUP) return 1;

            // 侧键 + 滚轮
            if (wParam == WM_MOUSEWHEEL) {
                isInteracted = true;
                short delta = (short)HIWORD(pMouse->mouseData);
                if (delta > 0) {
                    SendKey(g_config.sideWithWheelUp);
                }
                else {
                    SendKey(g_config.sideWithWheelDown);
                }
                return 1;
            }
        }
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

// 窗口过程函数
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    if (message == WM_TRAYICON) {
        if (lParam == WM_RBUTTONUP) {
            POINT curPoint;
            GetCursorPos(&curPoint);
            HMENU hMenu = CreatePopupMenu();

            UINT uEnableFlags = MF_STRING;
            if (g_isEnabled) uEnableFlags |= MF_CHECKED;
            AppendMenu(hMenu, uEnableFlags, ID_TRAY_TOGGLE, L"启用功能 Enable");
            AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);

            // 添加自启动选项并根据状态打钩
            UINT uFlags = MF_STRING;
            if (IsTaskRegistered()) uFlags |= MF_CHECKED;
            AppendMenu(hMenu, uFlags, ID_TRAY_AUTOSTART, L"开机自启动 Run at startup");

            AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
            AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, L"退出程序 Exit");

            SetForegroundWindow(hWnd);
            TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, curPoint.x, curPoint.y, 0, hWnd, NULL);
            DestroyMenu(hMenu);
        }
    }
    else if (message == WM_COMMAND) {
        int wmId = LOWORD(wParam);
        if (wmId == ID_TRAY_TOGGLE) {
            g_isEnabled = !g_isEnabled;
            // 如果关闭功能时侧键正被按下，重置状态以防按键粘连
            if (!g_isEnabled) {
                isSideButtonPressed = false;
                isInteracted = false;
            }
        }
        else if (wmId == ID_TRAY_AUTOSTART) {
            if (IsRunAsAdmin()) {
                bool currentState = IsTaskRegistered();
                ManageStartupTask(!currentState);
            }
            else {
                // 不是管理员，提权并在重启后自动执行
                ElevateNow();
            }
        }
        else if (wmId == ID_TRAY_EXIT) PostQuitMessage(0);
    }else if (message == g_uMsgTaskbarCreated && g_config.showTrayIcon) {
        // 重新调用 Shell_NotifyIcon(NIM_ADD, &nid)
        Shell_NotifyIcon(NIM_ADD, &nid);
        return 0;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}

// 主函数
int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    g_uMsgTaskbarCreated = RegisterWindowMessageW(L"TaskbarCreated");
    SetProcessDPIAware();
    if (wcsstr(GetCommandLineW(), L"--autostart")) {
        // 只有在确定已经是管理员权限时才执行
        if (IsRunAsAdmin()) {
            bool currentState = IsTaskRegistered();
            ManageStartupTask(!currentState);
        }
    }

    // 注册窗口类
    WNDCLASSEXW wcex = { sizeof(WNDCLASSEX) };
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.lpszClassName = L"MouseMediaControlClass";
    RegisterClassExW(&wcex);

    // 创建隐藏窗口
    HWND hWnd = CreateWindowExW(WS_EX_TOOLWINDOW, L"MouseMediaControlClass", L"", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    // 设置系统托盘图标
    if (g_config.showTrayIcon) {
        nid.cbSize = sizeof(NOTIFYICONDATA);
        nid.hWnd = hWnd;
        nid.uID = 1;
        nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        nid.uCallbackMessage = WM_TRAYICON;
        nid.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
        wcscpy_s(nid.szTip, g_config.trayTooltip);
        Shell_NotifyIcon(NIM_ADD, &nid);
    }

    // 安装鼠标钩子
    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, hInstance, 0);

    // 消息循环
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // 清理资源
    UnhookWindowsHookEx(hMouseHook);
    if (g_config.showTrayIcon) { Shell_NotifyIcon(NIM_DELETE, &nid); }

    return (int)msg.wParam;
}
