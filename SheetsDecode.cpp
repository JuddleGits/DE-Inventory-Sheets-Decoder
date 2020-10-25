// Creates a simple Win32 application that runs the SheetsDecode.py script when button is pressed
// Still a work in progress

#include <windows.h>
#include <Python.h>

void AddControls(HWND hwnd);
void ManageControls(HWND hwnd, WPARAM wParam);

const wchar_t g_szClassName[] = L"myWindowClass";

#define BUTTON 0x001

HWND button1;

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_CREATE:
		AddControls(hwnd);
		return 0;
	case WM_COMMAND:
		ManageControls(hwnd, wParam);
		return 0;
	case WM_CLOSE:
		DestroyWindow(hwnd);
		return 0;
	case WM_DESTROY:
		PostQuitMessage(0);
		return 0;
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
}

int WINAPI WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nCmdShow)
{
	WNDCLASSEX wc;
	HWND hwnd;
	MSG msg;

	wc.cbSize = sizeof(WNDCLASSEX);
	wc.style = 0;
	wc.lpfnWndProc = WndProc;
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = hInstance;
	wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = CreateSolidBrush(RGB(240, 240, 240));
	wc.lpszMenuName = NULL;
	wc.lpszClassName = g_szClassName;
	wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

	if (!(RegisterClassEx(&wc)))
	{
		MessageBox(NULL, L"Window Registration Failed", L"Error!",
			MB_ICONEXCLAMATION | MB_OK);
		return 0;
	}

	hwnd = CreateWindowEx(
		0,
		g_szClassName,
		L"Google Sheets Inventory Decoder",
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
		NULL, NULL, hInstance, NULL);

	if (hwnd == NULL)
	{
		MessageBox(NULL, L"Window Creation Failed!", L"Error!",
			MB_ICONEXCLAMATION | MB_OK);
	}

	if (hwnd != 0)
	{
		ShowWindow(hwnd, nCmdShow);
		UpdateWindow(hwnd);
	}
	else
		MessageBox(NULL, L"hwnd Equals 0", L"Error!",
			MB_ICONEXCLAMATION | MB_OK);

	while (GetMessage(&msg, NULL, 0, 0) > 0)
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return 0;
}

// Add default controls for first screen
void AddControls(HWND hwnd)
{
	button1 = CreateWindow(L"button", L"Decode Inventory Sheets",
		WS_VISIBLE | WS_CHILD,
		285, 250, 200, 30,
		hwnd, (HMENU)BUTTON, NULL, NULL);
}

// Manages actions taken when a button is pressed
void ManageControls(HWND hwnd, WPARAM wParam)
{
	switch (LOWORD(wParam))
	{
	case BUTTON:
		char fileName[] = "SheetsDecodePy.py";
		FILE* filePoint;
		Py_Initialize();
		filePoint = _Py_fopen(fileName, "r");
		PyRun_SimpleFile(filePoint, fileName);
		Py_Finalize();
		break;
	}
}
