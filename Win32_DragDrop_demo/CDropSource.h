/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#pragma once
#include "stdafx.h"


class CDropSource : public IDropSource
{

public:

	static HRESULT STDMETHODCALLTYPE CreateInstance(IDropSource **pDropSource);

	
	/*
	**  IUnknown **
	*/    
	HRESULT STDMETHODCALLTYPE QueryInterface (_In_ REFIID riid, _Out_ void   **ppvObject);
	ULONG   STDMETHODCALLTYPE AddRef();
	ULONG   STDMETHODCALLTYPE Release();

	
	/*
	** IDropSource **
	*/
	HRESULT STDMETHODCALLTYPE GiveFeedback(
		_In_ DWORD dwEffect
	);


	HRESULT STDMETHODCALLTYPE QueryContinueDrag(
		_In_ BOOL  fEscapePressed,
		_In_ DWORD grfKeyState
	);



	
private:

	//Private ctor and descrutor
	CDropSource();
	
	virtual  
	~CDropSource();


 UINT m_refCount;
 HWND m_hwnd;

 

};
