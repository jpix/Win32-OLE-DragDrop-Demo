/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
// Win32_DragDrop_demo.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "WinMain.h"

#define MAX_LOADSTRING 100



/*
Dialog based app funtion forward declaration
*/

INT_PTR CALLBACK MainDialogFunc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);


/***************************************************
WinMain entrypoint
****************************************************/
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.
	

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(INITCOMMONCONTROLSEX);

	
	InitCommonControlsEx(&icex);


	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_WIN32DLGPROJECT));

	//Main dlg window
	HWND MainDlg = CreateDialogParam(NULL, MAKEINTRESOURCE(IDD_MAIN_DIALOG), NULL, 
		                             reinterpret_cast<DLGPROC>(MainDialogFunc), NULL);
	ShowWindow(MainDlg, nCmdShow);
	UpdateWindow(MainDlg);


	MSG msg;

	
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		if (!IsDialogMessage(MainDlg, &msg))
		{
			
			if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
	}

	return (int)msg.wParam;


}





