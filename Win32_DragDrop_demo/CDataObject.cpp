/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#include "stdafx.h"
#include "CDataObject.h"




extern CStaticImage *cImage;
extern HBITMAP g_Hbm;
extern CString g_szText;



//Class activator
HRESULT CDataObject::CreateInstance(IDataObject ** pDataObject)
{
	HRESULT hr = E_OUTOFMEMORY;

	if (pDataObject != NULL) 
	{
		
		CDataObject *pCDataObj = new CDataObject();
		
		hr = pCDataObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID *>(pDataObject));
		
		if (SUCCEEDED(hr))
		return S_OK;
		
		
	}

	return hr;
}


//IUnknown implementation
HRESULT CDataObject::QueryInterface(REFIID riid, void ** ppvObject)
{
	
	if (IsEqualIID(riid, IID_IUnknown) || 
	    IsEqualIID(riid, IID_IDataObject))
		{

			*ppvObject = static_cast<IDataObject *>(this);

		}
	else
	return E_NOINTERFACE;
	

	AddRef();

	return S_OK;

}

ULONG CDataObject::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

ULONG CDataObject::Release()
{
	ULONG refCount = InterlockedDecrement(&m_refCount);
	if (refCount == 0)
	{
		delete this;
		return 0;
	}

	return refCount;
}




//IDataObject implementation

HRESULT CDataObject::DAdvise(FORMATETC * pformatetc, DWORD advf, IAdviseSink * pAdvSink, DWORD * pdwConnection)
{
	return E_NOTIMPL;
}

HRESULT CDataObject::DUnadvise(DWORD dwConnection)
{
	return E_NOTIMPL;
}

HRESULT CDataObject::EnumDAdvise(IEnumSTATDATA ** ppenumAdvise)
{
	return E_NOTIMPL;
}

HRESULT CDataObject::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC ** ppenumFormatEtc)
{
	
	HRESULT   hr;
	FORMATETC formatetc[] = {
	
	{ CF_TEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL | TYMED_ISTREAM },
	{ CF_UNICODETEXT, NULL, DVASPECT_CONTENT, -1,  TYMED_HGLOBAL | TYMED_ISTREAM },
	{ CF_BITMAP, NULL, DVASPECT_CONTENT, -1, TYMED_GDI }
	
};

	if (dwDirection == DATADIR_GET) {
		hr = SHCreateStdEnumFmtEtc(3, formatetc, ppenumFormatEtc);
		return hr;
	}
	else
		return E_NOTIMPL;
	


	
}

HRESULT CDataObject::GetCanonicalFormatEtc(FORMATETC * pformatectIn, FORMATETC * pformatetcOut)
{
	
	return E_NOTIMPL;
}

HRESULT CDataObject::GetData(FORMATETC * pformatetcIn, STGMEDIUM * pmedium)
{
	
	
	
	
	if(pformatetcIn->cfFormat == CF_TEXT ) {
		HGLOBAL hglobal;
		TCHAR    *p;

				
		hglobal = GlobalAlloc(GHND, m_szText.GetLength() + 1);

		p = (TCHAR *)GlobalLock(hglobal);
		lstrcpy(p, m_szText);
		GlobalUnlock(hglobal);

		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = hglobal;
		pmedium->pUnkForRelease = NULL;

		
	}
	
	else
		if (pformatetcIn->cfFormat == CF_UNICODETEXT) {
			HGLOBAL hglobal;
			LPWSTR    p;

		
			hglobal = GlobalAlloc(GHND, (m_szText.GetLength() + 1) * sizeof(TCHAR));

			p = (LPWSTR)GlobalLock(hglobal);
			lstrcpy(p, m_szText);
			GlobalUnlock(hglobal);

			pmedium->tymed = TYMED_HGLOBAL;
			pmedium->hGlobal = hglobal;
			pmedium->pUnkForRelease = NULL;

			return S_OK;
		}


	else 
		if (pformatetcIn->cfFormat == CF_BITMAP) 
		{
		
			
			if (NULL != cImage)			
				g_Hbm = cImage->Clone();

			
			

			pmedium->hBitmap = g_Hbm;
			pmedium->tymed = TYMED_GDI;
			pmedium->pUnkForRelease = NULL;
			
		
		}
	
	else
		return E_FAIL;
		

		return S_OK;


}

HRESULT CDataObject::QueryGetData(FORMATETC * pformatetc)
{
	if (pformatetc->cfFormat == CF_TEXT   || 
		pformatetc->cfFormat == CF_BITMAP || 
		pformatetc->cfFormat == CF_UNICODETEXT)
		return S_OK;

	return E_NOTIMPL;
}

HRESULT CDataObject::GetDataHere(FORMATETC * pformatetc, STGMEDIUM * pmedium)
{
		
	ULONG cb = (m_szText.GetLength() + 1) * sizeof(TCHAR);

	
	if ((pformatetc->cfFormat == CF_TEXT && pmedium->tymed == TYMED_ISTREAM)||
		(pformatetc->cfFormat == CF_UNICODETEXT && pmedium->tymed == TYMED_ISTREAM))
	{
		ULONG uWritten;
		pmedium->pstm->Write(m_szText, cb, &uWritten);
		uWritten = uWritten;
		
	}	
	else
		return E_FAIL;
		

	return S_OK;
}

HRESULT CDataObject::SetData(FORMATETC * pformatetc, STGMEDIUM * pmedium, BOOL fRelease)
{
	
	SetText(g_szText);

	return E_NOTIMPL;
}

void CDataObject::SetText(CString szText)
{

	m_szText = szText;

}



// default ctor / destructor
CDataObject::CDataObject()
{

	m_refCount = 0;

	m_szText = L"Text from OLE Drag&Drop";
}

CDataObject::~CDataObject()
{
}
