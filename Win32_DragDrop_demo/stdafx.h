/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"


#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h>

// C RunTime Header Files
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>


// TODO: reference additional headers your program requires here
//Common controls
#include <commctrl.h>

//Common dialogs
#include <Commdlg.h>

//OLE
#include <Ole2.h>
#include <OleIdl.h>

//ATL for CComPtr
#include <atlbase.h>
#include <atlstr.h>
#include <atlcoll.h>
#include <atlimage.h>

//Shell
#include <shlobj.h>
#include <shellapi.h>
#include <Shlwapi.h>

//Sal annotation
#include <sal.h>


#include "CStaticImage.h"
#include "CDropTarget.h"
#include "CDropSource.h"
#include "CDataObject.h"



#pragma comment(lib, "ComCtl32.lib")


#pragma comment(linker, \
  "\"/manifestdependency:type='Win32' "\
  "name='Microsoft.Windows.Common-Controls' "\
  "version='6.0.0.0' "\
  "processorArchitecture='*' "\
  "publicKeyToken='6595b64144ccf1df' "\
  "language='*'\"")
  

#define WM_LOADIMAGE WM_APP + 1
#define WM_UPDATE_GUI WM_APP + 2
#define DLG_MSG_CLEAR_DROPPEDFILES_LIST WM_APP + 3
#define DLG_MSG_GET_DROPPEDFILES_FROM_CLIPBOARD WM_APP + 4
#define DLG_MSG_SET_TEXT_FROM_CLIPBOARD WM_APP + 5
