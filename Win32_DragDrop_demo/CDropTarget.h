/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#pragma once
#include "stdafx.h"


class CDropTarget : public IDropTarget 
{

public:
	 HRESULT static CreateInstance(IDropTarget **pp);

protected:
	
	/*
	** IUnknown **
	*/
	HRESULT STDMETHODCALLTYPE QueryInterface(_In_ REFIID riid, _Out_ void   **ppvObject);
	ULONG   STDMETHODCALLTYPE AddRef();
	ULONG   STDMETHODCALLTYPE Release();
	

private:

	/*
	** IDropTarget ** 
	*/
	HRESULT STDMETHODCALLTYPE DragEnter(
		_In_      IDataObject *pDataObj,
		_In_      DWORD       grfKeyState,
		_In_      POINTL      pt,
		_Inout_ DWORD       *pdwEffect
	);

	
	 HRESULT STDMETHODCALLTYPE DragLeave();
	 	

	 HRESULT STDMETHODCALLTYPE DragOver(
		 _In_      DWORD  grfKeyState,
		 _In_      POINTL pt,
		 _Inout_ DWORD  *pdwEffect
	 );

	 
	 HRESULT STDMETHODCALLTYPE Drop(
		 _In_      IDataObject *pDataObj,
		 _In_      DWORD       grfKeyState,
		 _In_      POINTL      pt,
		 _Inout_   DWORD       *pdwEffect
	 );

	 

private:
	
	UINT m_refCount;
	
	CDropTarget();
	virtual 
		~CDropTarget();
	
	CComPtr<IDropTargetHelper> m_pDTH;
	HWND m_hwnd;

	CAtlList<CString> m_szFileList;

	


};
