/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#include "stdafx.h"

CStaticImage::CStaticImage()
{

	pImage = nullptr;

	
}

CStaticImage::CStaticImage(LPCTSTR fileName)
{
	
	m_Image.Load(fileName);


}

CStaticImage::CStaticImage(HBITMAP hBm)
{

	CImage tmpImage;
	IStream *pIstream = nullptr;
	HRESULT  hr = CreateStreamOnHGlobal(NULL, TRUE, &pIstream);

	tmpImage.Attach(hBm);

	hr = tmpImage.Save(pIstream, Gdiplus::ImageFormatBMP);


	m_Image.Load(pIstream);

	pIstream->Release();

}

CStaticImage::~CStaticImage()
{

	m_Image.Destroy();
	delete pImage;
	

}

HBITMAP CStaticImage::GetSafeHBITMAP()
{
	
	return static_cast<HBITMAP>(m_Image);
}





CStaticImage::CStaticImage(const CStaticImage &img)
{
	
	
	pImage = new CStaticImage();
	pImage->m_Image = img.m_Image;
		
	
}

int CStaticImage::Width()
{
	return m_Image.GetWidth();
}

int CStaticImage::Height()
{
	return m_Image.GetHeight();
}


BOOL CStaticImage::ScaleImage(int percent, SCALEPERCENT &scale)
{
	scale.cWidth = percent * Width() / 100;
	scale.cHeight = percent * Height() / 100;
	return TRUE;
}



CStaticImage CStaticImage::operator=(CStaticImage & obj)
{
	
	m_Image = obj.m_Image;
    return *this;
		
}

HBITMAP CStaticImage::GetImageCopy()
{
	return pImage->m_Image;
}

//Make a CImage copy
HBITMAP CStaticImage::Clone()
{
	
	//Make a CImage copy

	if (!Image.IsNull())
	{
		Image.Detach();
		Image.Destroy();
	}

	HBITMAP hBm = static_cast<HBITMAP>(m_Image);

	if (hBm != NULL)
	{
		DIBSECTION dbs;
		
		::GetObject(hBm, sizeof(DIBSECTION), &dbs);
		
		dbs.dsBmih.biCompression = BI_RGB;

		HDC hdc = ::GetDC(NULL);
		
		HBITMAP hbm_ddb = ::CreateDIBitmap(hdc, &dbs.dsBmih, CBM_INIT, 
			                                   dbs.dsBm.bmBits, (BITMAPINFO*)&dbs.dsBmih, 
			                                   DIB_RGB_COLORS);
		
		::ReleaseDC(NULL, hdc);

		return hbm_ddb;

	}

	DeleteObject(hBm);

	return hBm;
		
}

CImage CStaticImage::GetClonedImage()
{
	return Image;
}

HBITMAP CStaticImage::GetClonedSafeHBITMAP()
{
	return static_cast<HBITMAP>(Image);
}

VOID CStaticImage::SetImage(HBITMAP hBm)
{
	
	if (!m_Image.IsNull())
	{
		m_Image.Detach();
		m_Image.Destroy();
	}

	CImage tmpImage;
	IStream *pIstream = nullptr;
	HRESULT  hr = CreateStreamOnHGlobal(NULL, TRUE, &pIstream);

	tmpImage.Attach(hBm);

	hr = tmpImage.Save(pIstream, Gdiplus::ImageFormatBMP);

	
	m_Image.Load(pIstream);

	pIstream->Release();
	tmpImage.Detach();
	tmpImage.Destroy();

}

