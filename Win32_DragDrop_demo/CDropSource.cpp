/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#include "stdafx.h"
#include "CDropSource.h"

HRESULT CDropSource::CreateInstance(IDropSource **pDropSource)
{

	HRESULT hr = E_OUTOFMEMORY;


	if (pDropSource != NULL) 
	{
		
		CDropSource *pCDropSrc = new CDropSource();
		
		hr = pCDropSrc->QueryInterface(IID_IDropSource, reinterpret_cast<LPVOID *>(pDropSource));
		
		if (SUCCEEDED(hr))
			return S_OK;

		
	}

	return hr;


}

//IUnknown
HRESULT CDropSource::QueryInterface(REFIID riid, void ** ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDropSource))
	{
		*ppvObject = static_cast<IDropSource *>(this);
	}
	
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}

ULONG CDropSource::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

ULONG CDropSource::Release()
{
	ULONG refCount = InterlockedDecrement(&m_refCount);
	if (refCount == 0)
	{
		delete this;
		return 0;
	}

	return refCount;
}


//IDropSource
HRESULT CDropSource::GiveFeedback(DWORD dwEffect)
{
	
	
	return DRAGDROP_S_USEDEFAULTCURSORS;
}

HRESULT CDropSource::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
	if (fEscapePressed)
		return DRAGDROP_S_CANCEL;    
	
	if (!(grfKeyState & MK_LBUTTON))
		return DRAGDROP_S_DROP;         
	
	return S_OK;
}




// private Ctor and destructor
CDropSource::CDropSource()
{
	m_refCount = 0;
		
}

CDropSource::~CDropSource()
{
	
}
