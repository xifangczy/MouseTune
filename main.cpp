#include <windows.h>
#include <shellapi.h>
#include <taskschd.h>
#include <comdef.h>
#include <mmsystem.h>
#include "resource.h"

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "winmm.lib")

/// 应用程序配置结构
struct AppConfig {
    BYTE sideButton = XBUTTON2;          ///< 侧键（媒体控制主键）
    BYTE sideButtonAlone = VK_MEDIA_PLAY_PAUSE;  ///< 单独按下侧键
    BYTE sideWithLeft = VK_MEDIA_PREV_TRACK;     ///< 侧键 + 左键
    BYTE sideWithRight = VK_MEDIA_NEXT_TRACK;    ///< 侧键 + 右键
    BYTE sideWithWheelUp = VK_VOLUME_UP;         ///< 侧键 + 滚轮上
    BYTE sideWithWheelDown = VK_VOLUME_DOWN;     ///< 侧键 + 滚轮下

    BYTE scrollModeButton = XBUTTON1;    ///< 滚动模式触发按钮
    double scrollSensitivity = 0.8;      ///< 滚动灵敏度（每像素的滚轮增量倍数）
    bool reverseScroll = false;          ///< 是否反转滚动方向
    int minMoveThreshold = 2;            ///< 最小移动阈值（防抖）

    bool showTrayIcon = true;            ///< 是否显示托盘图标
    wchar_t trayTooltip[128] = L"MouseTune 鼠标侧键增强";  ///< 托盘提示文字
};

AppConfig g_config;  ///< 全局配置实例

/// 窗口消息定义
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_AUTOSTART 1002
#define ID_TRAY_TOGGLE 1003

const wchar_t* TASK_NAME = L"MouseTune_AutoStart";  ///< 计划任务名称
UINT g_uMsgTaskbarCreated = 0;  ///< 任务栏创建消息

bool g_isEnabled = true;        ///< 功能启用状态

HHOOK hMouseHook = nullptr;     ///< 鼠标钩子句柄
NOTIFYICONDATA nid = { 0 };     ///< 托盘图标数据

/// 侧键状态
bool isSideButtonPressed = false;  ///< 侧键是否按下
bool isInteracted = false;         ///< 是否已与其他键交互

/// 滚动模式状态
bool g_isScrollModeActive = false;   ///< 滚动模式是否激活
POINT g_scrollAnchor = { 0, 0 };     ///< 滚动锚点位置
HANDLE g_hScrollThread = nullptr;    ///< 滚动轮询线程句柄
volatile bool g_stopScrollThread = false;  ///< 停止滚动线程标志
bool scrollModePending = false;      ///< 滚动模式等待激活状态
long long scrollModePressTime = 0;       ///< 滚动按钮按下时间戳

/**
 * @brief 检查当前是否以管理员权限运行
 * @return true - 管理员权限，false - 普通权限
 */
bool IsRunAsAdmin() {
    BOOL fIsRunAsAdmin = FALSE;
    PSID AdministrationGroup;
    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;

    if (AllocateAndInitializeSid(&NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID,
        DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &AdministrationGroup)) {
        CheckTokenMembership(nullptr, AdministrationGroup, &fIsRunAsAdmin);
        FreeSid(AdministrationGroup);
    }
    return fIsRunAsAdmin == TRUE;
}

/**
 * @brief 请求管理员权限提升
 */
void ElevateNow() {
    wchar_t szPath[MAX_PATH];
    GetModuleFileNameW(nullptr, szPath, MAX_PATH);

    SHELLEXECUTEINFO sei = { sizeof(sei) };
    sei.lpVerb = L"runas";
    sei.lpFile = szPath;
    sei.lpParameters = L"--autostart";
    sei.hwnd = nullptr;
    sei.nShow = SW_NORMAL;

    if (ShellExecuteEx(&sei)) {
        PostQuitMessage(0);
    }
}

/**
 * @brief 管理开机自启动计划任务
 * @param enable true-创建任务，false-删除任务
 * @return 操作是否成功
 */
bool ManageStartupTask(bool enable) {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) return false;

    bool success = false;
    ITaskService* pService = nullptr;
    ITaskFolder* pRootFolder = nullptr;

    do {
        // 连接到计划任务服务
        if (FAILED(CoCreateInstance(CLSID_TaskScheduler, nullptr, CLSCTX_INPROC_SERVER,
            IID_ITaskService, (void**)&pService))) break;
        if (FAILED(pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()))) break;
        if (FAILED(pService->GetFolder(_bstr_t(L"\\"), &pRootFolder))) break;

		// 已经存在任务，删除它
        if (!enable) {
            pRootFolder->DeleteTask(_bstr_t(TASK_NAME), 0);
            success = true;
            break;
        }

        // 创建新任务
        ITaskDefinition* pTask = nullptr;
        if (FAILED(pService->NewTask(0, &pTask))) break;

        // 设置登录触发器
        ITriggerCollection* pTriggers = nullptr;
        pTask->get_Triggers(&pTriggers);
        ITrigger* pTrigger = nullptr;
        pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);
        pTriggers->Release();
        pTrigger->Release();

        // 设置执行动作
        IActionCollection* pActions = nullptr;
        pTask->get_Actions(&pActions);
        IAction* pAction = nullptr;
        pActions->Create(TASK_ACTION_EXEC, &pAction);
        IExecAction* pExecAction = nullptr;
        pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);

        wchar_t szPath[MAX_PATH];
        GetModuleFileNameW(nullptr, szPath, MAX_PATH);
        pExecAction->put_Path(_bstr_t(szPath));
        pExecAction->Release();
        pAction->Release();
        pActions->Release();

        // 设置管理员权限运行
        IPrincipal* pPrincipal = nullptr;
        pTask->get_Principal(&pPrincipal);
        pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
        pPrincipal->Release();

        // 设置登录延迟
        ILogonTrigger* pLogonTrigger = nullptr;
        pTrigger->QueryInterface(IID_ILogonTrigger, (void**)&pLogonTrigger);
        pLogonTrigger->put_Delay(_bstr_t(L"PT5S"));
        pLogonTrigger->Release();

        // 注册任务
        IRegisteredTask* pRegisteredTask = nullptr;
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

/**
 * @brief 检查计划任务是否已注册
 * @return true-已注册，false-未注册
 */
bool IsTaskRegistered() {
    bool found = false;
    ITaskService* pService = nullptr;
    ITaskFolder* pRootFolder = nullptr;

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);

    if (SUCCEEDED(CoCreateInstance(CLSID_TaskScheduler, nullptr, CLSCTX_INPROC_SERVER,
        IID_ITaskService, (void**)&pService))) {
        if (SUCCEEDED(pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()))) {
            if (SUCCEEDED(pService->GetFolder(_bstr_t(L"\\"), &pRootFolder))) {
                IRegisteredTask* pTask = nullptr;
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

/**
 * @brief 模拟按键按下和释放
 * @param vkey 虚拟键码
 */
void SendKey(BYTE vkey) {
    INPUT input = { 0 };
    input.type = INPUT_KEYBOARD;
    input.ki.wVk = vkey;
    SendInput(1, &input, sizeof(INPUT));

    input.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &input, sizeof(INPUT));
}

/**
 * @brief 模拟鼠标滚轮滚动
 * @param delta 滚轮增量（正数为向上，负数为向下）
 */
void SendWheel(int delta) {
    INPUT input = { 0 };
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_WHEEL;
    input.mi.mouseData = static_cast<DWORD>(delta);
    SendInput(1, &input, sizeof(INPUT));
}

/**
 * @brief 滚动轮询线程函数
 * @param lpParam 线程参数（未使用）
 * @return 线程退出码
 */
DWORD WINAPI ScrollPollingThread(LPVOID lpParam) {
    timeBeginPeriod(1);  // 提高定时器精度

    POINT lastPos = { 0, 0 };
    GetCursorPos(&lastPos);
    long long lastTime = GetTickCount64();

    while (!g_stopScrollThread) {
        Sleep(10);  // 100Hz轮询

        POINT curPos;
        if (!GetCursorPos(&curPos)) continue;

        long long now = GetTickCount64();
        if (now - lastTime < 8) continue;  // 限制更新频率

        int deltaY = curPos.y - lastPos.y;
        if (abs(deltaY) < g_config.minMoveThreshold) {
            lastPos = curPos;
            lastTime = now;
            continue;
        }

        // 计算滚轮增量
        double units = -deltaY * g_config.scrollSensitivity;
        if (g_config.reverseScroll) {
            units = -units;
        }

        int wheelDelta = static_cast<int>(units * WHEEL_DELTA);
        if (abs(wheelDelta) >= 10) {
            SendWheel(wheelDelta);
        }

        lastPos = curPos;
        lastTime = now;
    }

    timeEndPeriod(1);
    return 0;
}

/**
 * @brief 鼠标钩子回调函数
 * @param nCode 钩子代码
 * @param wParam 消息标识符
 * @param lParam 鼠标消息结构指针
 * @return 是否处理消息
 */
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    // 功能禁用时直接传递消息
    if (!g_isEnabled) {
        return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
    }

    if (nCode >= 0) {
        MSLLHOOKSTRUCT* pMouse = (MSLLHOOKSTRUCT*)lParam;

        // 忽略注入的事件
        if (pMouse->flags & LLMHF_INJECTED) {
            return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
        }

        // XBUTTON1: 滚动模式处理
        if (wParam == WM_XBUTTONDOWN && HIWORD(pMouse->mouseData) == g_config.scrollModeButton) {
            scrollModePending = true;
            scrollModePressTime = GetTickCount64();
            return 1;  // 阻止原始消息
        }
        else if (wParam == WM_XBUTTONUP && HIWORD(pMouse->mouseData) == g_config.scrollModeButton) {
            if (g_isScrollModeActive) {
                g_isScrollModeActive = false;
                g_stopScrollThread = true;

                if (g_hScrollThread) {
                    WaitForSingleObject(g_hScrollThread, 100);
                    CloseHandle(g_hScrollThread);
                    g_hScrollThread = nullptr;
                }
            }
            scrollModePending = false;
            return 1;  // 阻止原始消息
        }

        // 延迟100ms激活滚动模式
        if (scrollModePending && !g_isScrollModeActive) {
            long long now = GetTickCount64();
            if (now - scrollModePressTime >= 100) {
                scrollModePending = false;
                g_isScrollModeActive = true;
                g_stopScrollThread = false;
                g_hScrollThread = CreateThread(nullptr, 0, ScrollPollingThread, nullptr, 0, nullptr);
            }
        }

        // 媒体控制处理（滚动模式下不生效）
        if (!g_isScrollModeActive || g_config.sideButton != g_config.scrollModeButton) {
            // 侧键按下
            if (wParam == WM_XBUTTONDOWN && HIWORD(pMouse->mouseData) == g_config.sideButton) {
                isSideButtonPressed = true;
                isInteracted = false;
                return 1;  // 阻止原始消息
            }
            // 侧键释放
            else if (wParam == WM_XBUTTONUP && HIWORD(pMouse->mouseData) == g_config.sideButton) {
                isSideButtonPressed = false;
                if (!isInteracted) {
                    SendKey(g_config.sideButtonAlone);
                }
                return 1;  // 阻止原始消息
            }

            // 侧键组合功能
            if (isSideButtonPressed) {
                if (wParam == WM_LBUTTONDOWN) {
                    isInteracted = true;
                    SendKey(g_config.sideWithLeft);
                    return 1;
                }
                if (wParam == WM_LBUTTONUP) return 1;

                if (wParam == WM_RBUTTONDOWN) {
                    isInteracted = true;
                    SendKey(g_config.sideWithRight);
                    return 1;
                }
                if (wParam == WM_RBUTTONUP) return 1;

                if (wParam == WM_MOUSEWHEEL) {
                    isInteracted = true;
                    short zDelta = GET_WHEEL_DELTA_WPARAM(pMouse->mouseData);
                    SendKey(zDelta > 0 ? g_config.sideWithWheelUp : g_config.sideWithWheelDown);
                    return 1;
                }
            }
        }
    }

    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

/**
 * @brief 窗口过程函数
 * @param hWnd 窗口句柄
 * @param message 消息类型
 * @param wParam 消息参数1
 * @param lParam 消息参数2
 * @return 消息处理结果
 */
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    // 托盘图标消息处理
    if (message == WM_TRAYICON) {
        if (lParam == WM_RBUTTONUP) {
            POINT pt;
            GetCursorPos(&pt);
            HMENU hMenu = CreatePopupMenu();

            // 启用/禁用菜单项
            UINT uEnableFlags = MF_STRING;
            if (g_isEnabled) uEnableFlags |= MF_CHECKED;
            AppendMenu(hMenu, uEnableFlags, ID_TRAY_TOGGLE, L"启用功能");

            AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);

            // 自启动菜单项
            UINT autostartFlag = MF_STRING;
            if (IsTaskRegistered()) autostartFlag |= MF_CHECKED;
            AppendMenu(hMenu, autostartFlag, ID_TRAY_AUTOSTART, L"开机自启动 (管理员)");

            AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, L"退出");

            SetForegroundWindow(hWnd);
            TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, nullptr);
            DestroyMenu(hMenu);
        }
    }
    // 菜单命令处理
    else if (message == WM_COMMAND) {
        switch (LOWORD(wParam)) {
        case ID_TRAY_TOGGLE:
            g_isEnabled = !g_isEnabled;
            // 禁用时重置侧键状态
            if (!g_isEnabled) {
                isSideButtonPressed = false;
                isInteracted = false;
            }
            break;

        case ID_TRAY_AUTOSTART:
            if (IsRunAsAdmin()) {
                ManageStartupTask(!IsTaskRegistered());
            }
            else {
                ElevateNow();
            }
            break;

        case ID_TRAY_EXIT:
            PostQuitMessage(0);
            break;
        }
    }
    // 任务栏重建后恢复托盘图标
    else if (message == g_uMsgTaskbarCreated && g_config.showTrayIcon) {
        Shell_NotifyIcon(NIM_ADD, &nid);
    }

    return DefWindowProc(hWnd, message, wParam, lParam);
}

/**
 * @brief 应用程序主函数
 * @param hInstance 当前实例句柄
 * @param hPrevInstance 前一个实例句柄（已废弃）
 * @param lpCmdLine 命令行参数
 * @param nCmdShow 窗口显示方式
 * @return 应用程序退出码
 */
int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    // 自启动模式（无界面）
    if (wcsstr(GetCommandLineW(), L"--autostart")) {
        if (IsRunAsAdmin()) {
            ManageStartupTask(true);
        }
        return 0;
    }

    SetProcessDPIAware();  // 启用DPI感知
    g_uMsgTaskbarCreated = RegisterWindowMessage(L"TaskbarCreated");

    // 注册窗口类
    WNDCLASSEX wc = { sizeof(WNDCLASSEX) };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"MouseTuneClass";
    RegisterClassEx(&wc);

    // 创建隐藏窗口
    HWND g_hWnd = CreateWindowEx(WS_EX_TOOLWINDOW, L"MouseTuneClass", L"",
        0, 0, 0, 0, 0, nullptr, nullptr, hInstance, nullptr);

    // 设置托盘图标
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = g_hWnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
    wcscpy_s(nid.szTip, g_config.trayTooltip);
    Shell_NotifyIcon(NIM_ADD, &nid);

    // 安装鼠标钩子
    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, hInstance, 0);
    if (!hMouseHook) {
        MessageBox(nullptr, L"安装鼠标钩子失败！", L"错误", MB_ICONERROR);
        return 1;
    }

    // 消息循环
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // 清理资源
    if (g_isScrollModeActive) {
        g_stopScrollThread = true;
        if (g_hScrollThread) {
            TerminateThread(g_hScrollThread, 0);
            CloseHandle(g_hScrollThread);
        }
    }

    UnhookWindowsHookEx(hMouseHook);
    Shell_NotifyIcon(NIM_DELETE, &nid);
    DestroyIcon(nid.hIcon);

    return static_cast<int>(msg.wParam);
}
