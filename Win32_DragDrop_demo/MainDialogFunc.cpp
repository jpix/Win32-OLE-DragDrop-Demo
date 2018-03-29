/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#include "stdafx.h"
#include "MainDialogFunc.h"

/*
** Global variables declaration **
*/

HWND g_hwnd_Dlg;
CStaticImage *cImage = nullptr;
CAtlList<CString> DroppedFiles;
HBITMAP g_Hbm = nullptr;
CString g_szText;


/*Drag&Drop Source and IDataObject*/
IDropSource *pDs;
IDataObject *pDo;


/*Drag&Drop Target*/
IDropTarget *pDt;

/* ** ** ** ** ** ** */

/*
** Dialog window procedure **
*/
INT_PTR CALLBACK MainDialogFunc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	
	HBRUSH hbrBkgnd = nullptr;

	switch (uMsg)
	{

	case WM_NCCREATE:
	{

		if ((BOOL)(((LPCREATESTRUCT)lParam)->lpCreateParams))
			EnableNonClientDpiScaling(hDlg);

		return DefWindowProc(hDlg, uMsg, wParam, lParam);

	}


	case WM_INITDIALOG:

	{
		
		OleInitialize(NULL);
		
		
		UINT uDpi = 96;

		// Determine the DPI to use, accoridng to the DPI awareness mode
		DPI_AWARENESS dpiAwareness = GetAwarenessFromDpiAwarenessContext(GetThreadDpiAwarenessContext());
		switch (dpiAwareness)
		{
			// Scale the window to the system DPI
		case DPI_AWARENESS_SYSTEM_AWARE:
			uDpi = GetDpiForSystem();
			break;

			// Scale the window to the monitor DPI
		case DPI_AWARENESS_PER_MONITOR_AWARE:
			uDpi = GetDpiForWindow(hDlg);
			break;

		case DPI_AWARENESS_UNAWARE:

			break;
		}

		/*
		*** Set font for DPI ***
		*/
		
		LOGFONT lfText = {};
		SystemParametersInfoForDpi(SPI_GETICONTITLELOGFONT, sizeof(lfText), &lfText, FALSE, uDpi);
		HFONT hFontNew = CreateFontIndirect(&lfText);
		SendMessage(hDlg, WM_SETFONT, (WPARAM)hFontNew, MAKELPARAM(TRUE, 0));
		
		
		g_hwnd_Dlg = hDlg;
		Button_SetCheck(GetDlgItem(hDlg, IDC_CHECK1),TRUE);
		
		/*
		*** DRAG AND DROP TARGET *** 
		*/
		HRESULT hr = E_OUTOFMEMORY;

		hr = CDropTarget::CreateInstance(&pDt);
		hr = RegisterDragDrop(hDlg, pDt);
		hr = pDt->Release();


		/*
		*** DRAG AND DROP SOURCE ***
		*/
		hr = CDropSource::CreateInstance(&pDs);
		
		
		/*
		*** IDataObject ***
		*/
		hr = CDataObject::CreateInstance(&pDo);

		
	}
	break;

	
	case WM_CTLCOLORDLG:

		/*
		*** Dlg background color ***
		*/
	
	return (INT_PTR)GetStockObject(WHITE_BRUSH);
		break;

		
		
	case WM_CTLCOLORSTATIC:
	{

		/*
		** Set controls background **
		*/
		HWND hwndCtrl = (HWND)lParam;
		HDC hdcStatic = (HDC)wParam;
		SetTextColor(hdcStatic, RGB(0, 0, 0));
		SetBkColor(hdcStatic, RGB(255, 255, 255));

		if (hwndCtrl == GetDlgItem(hDlg, IDC_STATIC1) || 
			hwndCtrl == GetDlgItem(hDlg, IDC_STATIC2) ||
			hwndCtrl == GetDlgItem(hDlg, IDC_CHECK1))
		
		{

			if (hbrBkgnd == NULL)
			{
				hbrBkgnd = CreateSolidBrush(RGB(255, 255, 255));
			}
			return (INT_PTR)hbrBkgnd;
		}
		
	

	}
	break;
	
	
	
	case WM_PAINT:
	{

		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(hDlg, &ps);
		
		// TODO: Add any drawing code that uses hdc here...
		
		EndPaint(hDlg, &ps);
	}
	break;

	



	case WM_LOADIMAGE:
	{
	
		ShowWindow(GetDlgItem(hDlg, IDC_IMAGE), TRUE);
		UpdateWindow(GetDlgItem(hDlg, IDC_IMAGE));

		if (NULL != cImage)
			delete cImage;
		

		cImage = new CStaticImage(L"d:\\com_dragdrop.png");
				
				
		/*
		*** Create CImage Copy
		*/
		/*
		HBITMAP hh = cImage->Clone();

		SendMessage(GetDlgItem(hDlg, IDC_IMAGE), STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)(HBITMAP)hh);

		DeleteObject(hh);
		*/


		SendMessage(hDlg, WM_UPDATE_GUI, (WPARAM)NULL, (LPARAM)NULL);
		
		
		return TRUE;
		

	}
		break;



	case DLG_MSG_CLEAR_DROPPEDFILES_LIST:
	{
		DroppedFiles.RemoveAll();

		break;
	}

	case DLG_MSG_SET_TEXT_FROM_CLIPBOARD:
	{

		CString szText = CString((LPWSTR)lParam);

		Edit_SetText(GetDlgItem(hDlg, IDC_EDIT1), szText.GetBuffer());

		szText.ReleaseBuffer();

		break;
	}



	case DLG_MSG_GET_DROPPEDFILES_FROM_CLIPBOARD:
	{


		DroppedFiles.AddTailList((CAtlList<CString>*)(lParam));

		int count = DroppedFiles.GetCount();
		CString szText;

		POSITION pos = DroppedFiles.GetHeadPosition();


		TCHAR buff[MAX_PATH] = { 0 };
		Edit_GetText(GetDlgItem(hDlg, IDC_EDIT1), buff, MAX_PATH);

		CString fileName = DroppedFiles.GetAt(pos);
		CString extension = PathFindExtension(DroppedFiles.GetAt(pos));
		if (extension.Compare(L".png")  == 0 ||
			extension.Compare(L".jpg")  == 0 ||
			extension.Compare(L".jpeg") == 0 ||
			extension.Compare(L".bmp")  == 0)

		{
			
			if(NULL != cImage)
			delete cImage;
			
			cImage = new CStaticImage(fileName);
			SendMessage(hDlg, WM_UPDATE_GUI, NULL, NULL);
			
		}
	
		
		szText = buff;

		
		do {
						
			szText += DroppedFiles.GetNext(pos);
			szText += L"\r\n";
			
		} while (pos != NULL);

		Edit_SetText(GetDlgItem(hDlg, IDC_EDIT1), szText);

		break;
	}





		case WM_UPDATE_GUI:
		{
			

						
			
			ShowWindow(GetDlgItem(hDlg, IDC_IMAGE), SW_SHOW);
			UpdateWindow(GetDlgItem(hDlg, IDC_IMAGE));

			
			
			g_Hbm = reinterpret_cast<HBITMAP>(lParam);

				
				
				
				if (g_Hbm != NULL)
				{

					if (NULL != cImage)
						delete cImage;


					cImage = new CStaticImage(g_Hbm);
				
			    }
			
						
	
					RECT rct;
					GetClientRect(GetDlgItem(hDlg, IDC_STATIC1),&rct);

					int maxWidth = rct.right - rct.left;
					int maxHeight = rct.bottom - rct.top;
					
					CStaticImage::SCALEPERCENT scale;

					
					if (cImage->Width() > (maxWidth) || cImage->Height() > (maxHeight))
					{
						
						cImage->ScaleImage(60, scale);

					}

					else
						cImage->ScaleImage(100, scale);
					

					
					int width = scale.cWidth;
					int height = scale.cHeight;

					GetClientRect(hDlg, &rct);
					int cX = rct.right + rct.left;
					int cY = rct.bottom - rct.top;

					SendMessage(GetDlgItem(hDlg, IDC_IMAGE), STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)cImage->GetSafeHBITMAP());

					MoveWindow(GetDlgItem(hDlg, IDC_IMAGE), (cX / 2) - (width / 2),  50 , width, height, TRUE);
										
		
			
		}

			break;
		

	case WM_SIZE:
	{


		SendMessage(hDlg, WM_UPDATE_GUI, (WPARAM)NULL, (LPARAM)NULL);


	}
		break;




	case WM_LBUTTONDOWN:
	{

	
		if (NULL == cImage)
			return TRUE;
		
		HRESULT hr = E_FAIL;

		BOOL checked = Button_GetCheck(GetDlgItem(hDlg, IDC_CHECK1));
		
		if (checked)
		{
			TCHAR szString[MAX_PATH];
			Edit_GetText(GetDlgItem(hDlg, IDC_EDIT1), szString, MAX_PATH);
			g_szText = szString;
			pDo->SetData(NULL, NULL, NULL);
		}

		hr = SHCreateDataObject(NULL, NULL, NULL, pDo, IID_IDataObject, reinterpret_cast<LPVOID*>(&pDo));

		hr = OleSetClipboard(pDo);

	 	DWORD dwEffect;
		
		//Set CDataSource pointer pDs to NULL if you want to use the Shell DropSource class
		hr = SHDoDragDrop(hDlg, pDo, pDs /*NULL*/, DROPEFFECT_COPY, &dwEffect);

	}

	break;

	case WM_COMMAND:
	{
		int wmId = LOWORD(wParam);
		
		switch (wmId)
		{
		
			/*
			*** Button commands
			*/

	
		case IDCANCEL:
			SendMessage(hDlg, WM_CLOSE, 0, 0);
			return TRUE;

		case IDOK:
		{
			
			SendMessage(hDlg, WM_COMMAND, (WPARAM)IDM_IMAGE_LOAD, (LPARAM)(NULL));
			
			return TRUE;
			
		}

		/****************************************/

		/*
		*** Menu commands
		*/

		case IDM_EXIT:
			DestroyWindow(hDlg);
			break;


		case IDM_ABOUT:
			DialogBox(NULL, MAKEINTRESOURCE(IDD_ABOUTBOX), hDlg, About);
			break;


		case IDM_IMAGE_LOAD:
		{
			OPENFILENAMEW ofn;
			TCHAR szFileName[MAX_PATH] = L"";

			ZeroMemory(&ofn, sizeof(ofn));

			ofn.lStructSize = sizeof(ofn);
			ofn.hwndOwner = hDlg;
			ofn.lpstrFilter = L"Images (*.png;*.jpg;*.jpeg;*.bmp)\0*.png;*.jpg;*.jpeg;*.bmp\0PNG files (*.png)\0*.png\0JPG files (*.jpg,*jpeg)\0*.jpg\0BMP files (*.bmp)\0*.bmp\0\0";
			ofn.lpstrFile = szFileName;
			ofn.nMaxFile = MAX_PATH;
			ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;


			if (GetOpenFileName(&ofn))
			{
				CString fileName = ofn.lpstrFile;
				CString extension = PathFindExtension(fileName);

				//LPCTSTR sz_filename = fileName;

				ShowWindow(GetDlgItem(hDlg, IDC_IMAGE), TRUE);
				UpdateWindow(GetDlgItem(hDlg, IDC_IMAGE));

				if (NULL != cImage)
					delete cImage;


				cImage = new CStaticImage(fileName);

			   SendMessage(hDlg, WM_UPDATE_GUI, (WPARAM)(NULL), (LPARAM)(NULL));

			

			}
		}
			
		break;




		case IDM_IMAGE_CLEAR:
		{

			SendMessage(GetDlgItem(hDlg, IDC_IMAGE), STM_SETIMAGE, (WPARAM)IMAGE_BITMAP,(LPARAM)(NULL));
		}
		break;


		case IDM_CLIPBOARD_CLEAR:
		{
		
			OpenClipboard(hDlg);
			EmptyClipboard();
			CloseClipboard();

		
		
		}
			break;




		case IDM_IMAGE_COPYTOCLIPBOARD:
		{

			if (NULL == cImage)
				break;

			HBITMAP hBm = cImage->Clone();

			if (NULL != hBm)
			{
				OpenClipboard(NULL);
				EmptyClipboard();
				SetClipboardData(CF_BITMAP, hBm);
				CloseClipboard();
			}
			DeleteObject(hBm);
		}
		break;

		/************************/	
		

		default:
			return DefWindowProc(hDlg, uMsg, wParam, lParam);
		}
	}
	break;


	case WM_CLOSE:
		
	{
		DestroyWindow(hDlg);
	}
	return TRUE;

	case WM_DESTROY:
		OleUninitialize();
		delete cImage;
		RevokeDragDrop(hDlg);
		DeleteObject(hbrBkgnd);
		PostQuitMessage(0);
		return TRUE;
	}

	return FALSE;
}




// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
