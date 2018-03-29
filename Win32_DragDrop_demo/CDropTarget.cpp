/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#include "stdafx.h"
#include "CDropTarget.h"



extern HWND g_hwnd_Dlg;
extern HBITMAP g_Hbm;

HRESULT CDropTarget::CreateInstance(IDropTarget ** pDropTarget)
{
	HRESULT hr = E_OUTOFMEMORY;
	
	if (pDropTarget != NULL) 
	{
	
		CDropTarget *pCDropTarget = new CDropTarget();
		
		hr = pCDropTarget->QueryInterface(IID_IDropTarget, reinterpret_cast<LPVOID *>(pDropTarget));
				
		if (SUCCEEDED(hr))
		return S_OK;
	}

	return hr;

}

//IUnknown
HRESULT CDropTarget::QueryInterface(REFIID riid, void ** ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDropTarget))
		*ppvObject = static_cast<IDropTarget *>(this);
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}

ULONG CDropTarget::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

ULONG CDropTarget::Release()
{
	ULONG refCount = InterlockedDecrement(&m_refCount);
	if (refCount == 0)
	{
		delete this;
		return 0;
	}

	return refCount;
}

//IDropTarget
HRESULT CDropTarget::DragEnter(IDataObject * pDataObj, DWORD grfKeyState, POINTL pt, DWORD * pdwEffect)
{
	if (m_pDTH) 
	{
		POINT pts = { pt.x, pt.y };
		
		m_pDTH->DragEnter(m_hwnd, pDataObj, &pts, *pdwEffect);
	}

	return S_OK;
}

HRESULT CDropTarget::DragLeave()
{
	m_pDTH->DragLeave();

	return S_OK;
}

HRESULT CDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD * pdwEffect)
{
	if (m_pDTH) 
	{
		POINT pts = { pt.x, pt.y };
		
		m_pDTH->DragOver(&pts, *pdwEffect);
	}

	return S_OK;
}

HRESULT CDropTarget::Drop(IDataObject * pDataObj, DWORD grfKeyState, POINTL pt, DWORD * pdwEffect)
{
	if (m_pDTH) 
	{
		POINT ptx = { pt.x, pt.y };
		
		m_pDTH->Drop(pDataObj, &ptx, *pdwEffect);
	}

		
	
	FORMATETC  txt_formatetc;
    txt_formatetc.cfFormat = CF_UNICODETEXT;
	txt_formatetc.ptd = NULL;
	txt_formatetc.dwAspect = DVASPECT_CONTENT;
	txt_formatetc.lindex = -1;
	txt_formatetc.tymed = TYMED_HGLOBAL;
	

	FORMATETC  dropfiles_formatetc;
	dropfiles_formatetc.cfFormat = CF_HDROP;
	dropfiles_formatetc.ptd = NULL;
	dropfiles_formatetc.dwAspect = DVASPECT_CONTENT;
	dropfiles_formatetc.lindex = -1;
	dropfiles_formatetc.tymed = TYMED_HGLOBAL;


	FORMATETC  bitmap_formatetc;
	bitmap_formatetc.cfFormat = CF_BITMAP;
	bitmap_formatetc.ptd = NULL;
	bitmap_formatetc.dwAspect = DVASPECT_CONTENT;
	bitmap_formatetc.lindex = -1;
	bitmap_formatetc.tymed = TYMED_GDI;

	

	STGMEDIUM medium;

	
	HRESULT hr = E_OUTOFMEMORY;

	// check if formatetc is valid for text
	hr = pDataObj->QueryGetData(&txt_formatetc);
	
	if (SUCCEEDED(hr)) //formatetc is valid for text
	{
		hr = pDataObj->GetData(&txt_formatetc, &medium);

		if (SUCCEEDED(hr)) 

		{
			
			LPWSTR pData = (LPWSTR) ::GlobalLock(medium.hGlobal);

			

			SendMessage(g_hwnd_Dlg, DLG_MSG_SET_TEXT_FROM_CLIPBOARD, (WPARAM)(NULL), (LPARAM)(pData));
			
			
		}

		GlobalUnlock(medium.hGlobal);
		ReleaseStgMedium(&medium);


	}

 
	// check if formatetc is valid for HDROP
	hr = pDataObj->QueryGetData(&dropfiles_formatetc);
  
	if (SUCCEEDED(hr)) // formatetc is valid for HDDROP
	{
		
		m_szFileList.RemoveAll();
		SendMessage(g_hwnd_Dlg, DLG_MSG_CLEAR_DROPPEDFILES_LIST, (WPARAM)(NULL), (LPARAM)(NULL));
		
		hr = pDataObj->GetData(&dropfiles_formatetc, &medium);

		if (SUCCEEDED(hr)) 
		{

			HDROP p = static_cast<HDROP>(GlobalLock(medium.hGlobal));


			int num = DragQueryFile(p, 0xFFFFFFFF, NULL, 0);
			for (int iFile = 0; iFile < num; iFile++) {
				int size = DragQueryFile(p, iFile, NULL, 0);

				TCHAR * buf = new TCHAR[size + 1];


				DragQueryFile(p, iFile, buf, size + 1);


				m_szFileList.AddTail(CString(buf));

				delete[] buf;
			}


			SendMessage(g_hwnd_Dlg, DLG_MSG_GET_DROPPEDFILES_FROM_CLIPBOARD, (WPARAM)(NULL), (LPARAM)(&m_szFileList));

			GlobalUnlock(medium.hGlobal);
			ReleaseStgMedium(&medium);

		}
	}


	// check if formatetc is valid for images
	hr = pDataObj->QueryGetData(&bitmap_formatetc);

	if (SUCCEEDED(hr)) // formatetc is valid for images
	{
		
		

		hr = pDataObj->GetData(&bitmap_formatetc, &medium);

		if (SUCCEEDED(hr)) 
		{

			
			if (medium.hBitmap != NULL)
			{

				SendMessage(g_hwnd_Dlg, WM_UPDATE_GUI, (WPARAM)(NULL), (LPARAM)medium.hBitmap);
				
			}
			
			GlobalUnlock(medium.hBitmap);
			ReleaseStgMedium(&medium);

			
		}



	}

		
	if ((grfKeyState & MK_CONTROL) == MK_CONTROL)
	{
		*pdwEffect = DROPEFFECT_COPY;
	}
	else
		if ((grfKeyState & MK_ALT) == MK_ALT)
		{
			*pdwEffect = DROPEFFECT_MOVE;
		}
		else

	
		m_pDTH.p->Release();
	

	return S_OK;


}


//Default ctor
CDropTarget::CDropTarget()
{
	m_refCount = 0;

	CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
		IID_IDropTargetHelper, (LPVOID*)&m_pDTH);
		

}


CDropTarget::~CDropTarget()
{

}
