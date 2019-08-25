#include "stdafx.h"
#include <shellapi.h>
#include "resource.h"
#include "events.h"
#include "debugtrace.h"

//globals
WNDCLASSEX g_wcex; //reuse g_wcex.hInstance and g_wcex.hIcon
#define TRAY_WNDCLASS_NAME  _T("IEmaximizerTray")
HMENU g_htrayMenu;
NOTIFYICONDATA  g_trayicon;
#define WM_TRAY_NOTIFY (WM_APP+11)

HWINEVENTHOOK g_hWinEventHook;
BOOL g_bPopup;

#define APPWNDCLASSNAME  _T("IEFrame")

//https://docs.microsoft.com/en-us/windows/win32/winauto/event-constants
const DWORD g_arrWinEvents[]= {
	EVENT_OBJECT_CREATE,
	//EVENT_OBJECT_DESTROY,
	EVENT_OBJECT_SHOW
	//EVENT_OBJECT_HIDE,
	//EVENT_OBJECT_REORDER,

	//EVENT_SYSTEM_FOREGROUND,
	//EVENT_OBJECT_LOCATIONCHANGE,
	//EVENT_OBJECT_STATECHANGE
};

//https://solarianprogrammer.com/2016/11/28/cpp-passing-c-style-array-with-size-information-to-function/
template <class T, size_t N> T GetArrayMin(const T (&arr)[N])
{
	T m = arr[0];
	size_t i;
	for (i = 1; i < N; i++)
	{
		if (arr[i] < m) m = arr[i];
	}
return m;
};
template <class T, size_t N> T GetArrayMax(const T (&arr)[N])
{
	T m = arr[0];
	size_t i;
	for (i = 1; i < N; i++)
	{
		if (arr[i] > m) m = arr[i];
	}
return m;
};

BOOL IsAppWindowClass(HWND hw)
{
return ((GetClassName(hw, g_trayicon.szTip, 128)==(_countof(APPWNDCLASSNAME)-1) ) && (lstrcmp(g_trayicon.szTip, APPWNDCLASSNAME)==0));
}

/////////////////////////////////////////////////////////////
void CALLBACK WinEventProcCallback(HWINEVENTHOOK hook, DWORD dwEvent, HWND hwnd, LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime)
{
	//DBGTRACEEVENT(dwEvent);
	if ((idObject == OBJID_WINDOW) && (idChild == CHILDID_SELF))
		switch (dwEvent)
		{
		case EVENT_OBJECT_CREATE:
			{
				if (IsAppWindowClass(hwnd))
				{
					DBGTRACE("EVENT_OBJECT_CREATE 0x%x\n", hwnd);
					//ignore IE popup windows (eg. when viewing single message in Outlook Web App)
					g_bPopup = IsAppWindowClass(GetForegroundWindow());
				}
			}
			break;
		//earliest event that won't change our custom size/position
		case EVENT_OBJECT_SHOW:
			{
				if (IsAppWindowClass(hwnd))
				{
					if (!g_bPopup)
					{
						g_bPopup = FALSE;
						DBGTRACE("EVENT_OBJECT_SHOW  0x%x\n", hwnd);
						PostMessage(hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, NULL);
					}
				}
			}
			break;

			//#ifdef _DEBUG
			//		case EVENT_SYSTEM_FOREGROUND:
			//if (IsAppWindowClass(hwnd)) DBGTRACE("EVENT_SYSTEM_FOREGROUND\n");
			//			break;
			//		case EVENT_OBJECT_HIDE:
			//if (IsAppWindowClass(hwnd)) DBGTRACE("EVENT_OBJECT_HIDE\n");
			//			break;
			//		case EVENT_OBJECT_DESTROY:
			//if (IsAppWindowClass(hwnd)) DBGTRACE("EVENT_OBJECT_DESTROY\n");
			//			break;
			//		case EVENT_OBJECT_LOCATIONCHANGE:
			//if (IsAppWindowClass(hwnd)) DBGTRACE("EVENT_OBJECT_LOCATIONCHANGE\n");
			//			break;
			//		case EVENT_OBJECT_STATECHANGE:
			//if (IsAppWindowClass(hwnd)) DBGTRACE("EVENT_OBJECT_STATECHANGE\n");
			//			break;
			//		case EVENT_OBJECT_REORDER:
			//if (IsAppWindowClass(hwnd)) DBGTRACE("EVENT_OBJECT_REORDER\n");
			//			break;
			//#endif
		}
}

// Initialize COM and set up the event hook
void InitializeMSAA()
{
	if (g_hWinEventHook==NULL)
	{
		//DBGTRACE("GetArrayMin=0x%x  GetArrayMax=0x%x\n",GetArrayMin(g_arrWinEvents), GetArrayMax(g_arrWinEvents));
		g_hWinEventHook = SetWinEventHook(GetArrayMin(g_arrWinEvents),
										GetArrayMax(g_arrWinEvents), //EVENT_MIN, EVENT_MAX /*EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND*/,
											NULL,  // handle to DLL
											WinEventProcCallback,
											0, 0, // process and thread IDs of interest (0 = all)
											WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
	}
}
// Unhook the events and shut down COM
void ShutdownMSAA()
{
    if (g_hWinEventHook)
	{
		UnhookWinEvent(g_hWinEventHook);
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//message-only window proc
LRESULT CALLBACK TrayWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_COMMAND://menu Exit command
				DestroyWindow(hWnd);
		break;

		case WM_TRAY_NOTIFY:
			if (LOWORD(lParam)==WM_RBUTTONUP)
			{
				POINT mousepos;
				if (!GetCursorPos(&mousepos)) return 0;
				::SetForegroundWindow(hWnd);
				::TrackPopupMenuEx(GetSubMenu(g_htrayMenu, 0), TPM_RIGHTALIGN, mousepos.x, mousepos.y, hWnd, NULL);
				::PostMessage(hWnd, WM_NULL, 0, 0);
			}
		break;

		case WM_ENDSESSION:
			if (wParam==TRUE) { ShutdownMSAA(); DestroyWindow(hWnd);}
		break;

		case WM_DESTROY:
			PostQuitMessage(0);
		break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
	}
return 0;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//main
#ifdef _ATL_MIN_CRT //release build
int WINAPI WinMainCRTStartup(void)//no crt
#else
int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
#endif
{
	CoInitialize(NULL);

	//init globals
	SecureZeroMemory(&g_wcex, sizeof(WNDCLASSEX));
	g_htrayMenu = NULL;
	SecureZeroMemory(&g_trayicon, sizeof(NOTIFYICONDATA));
	g_hWinEventHook = NULL;
	g_bPopup = FALSE;

	//prep
	g_wcex.cbSize			= sizeof(WNDCLASSEX);
	g_wcex.hInstance		= GetModuleHandle(NULL);
	g_wcex.lpfnWndProc	= TrayWndProc;
	g_wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_MAIN);
	g_wcex.lpszClassName	= TRAY_WNDCLASS_NAME;
	g_wcex.hIconSm = (HICON)::LoadImage(g_wcex.hInstance, MAKEINTRESOURCE(IDI_MAIN),
					IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);

	if (!RegisterClassEx(&g_wcex)) return 0;

	//create message-only window
	g_trayicon.hWnd = CreateWindow(g_wcex.lpszClassName, NULL, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT,
									0, CW_USEDEFAULT, 0, HWND_MESSAGE, NULL, g_wcex.hInstance, NULL);
	if (!g_trayicon.hWnd) return FALSE;

	g_htrayMenu = ::GetMenu(g_trayicon.hWnd);
	
	//MessageBox(NULL, GetCommandLine(), NULL,MB_OK);
	if (_tcsstr(GetCommandLine(), _T("/notray")) == NULL)
	{
		//fill tray icon info
		g_trayicon.cbSize = sizeof(NOTIFYICONDATA);
		g_trayicon.uFlags = (NIF_ICON | NIF_MESSAGE | NIF_TIP);
		//g_trayicon.hWnd already set
		g_trayicon.hIcon = g_wcex.hIconSm;
		g_trayicon.dwInfoFlags  =NIIF_USER;
		g_trayicon.uCallbackMessage = WM_TRAY_NOTIFY;
		g_trayicon.uTimeout = 3000;//ms
		LoadString(GetModuleHandle(NULL), IDS_TRAYTOOLTIP, g_trayicon.szTip, _countof(g_trayicon.szTip));
		//add tray icon
		Shell_NotifyIcon(NIM_ADD, &g_trayicon);	
	}

	InitializeMSAA();

	// Main message loop
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	ShutdownMSAA();

	//cleanup
	if (g_trayicon.cbSize) Shell_NotifyIcon(NIM_DELETE, &g_trayicon);
	
	CoUninitialize();
#ifdef _ATL_MIN_CRT
	ExitProcess(0);
#endif
return 0;
}
