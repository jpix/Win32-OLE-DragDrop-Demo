/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#pragma once

#include "stdafx.h"
#include "resource.h"

/*
Dialog based app funtion forward declaration
*/

/*
Main dlg function
*/
INT_PTR CALLBACK MainDialogFunc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);


/*
Aboutbox dlg function
*/
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);


