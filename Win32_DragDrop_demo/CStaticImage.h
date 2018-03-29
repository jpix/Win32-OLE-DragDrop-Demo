/* ** ** ** ** ** ** ** ** ** ** **

Win32 OLE Drag&Drop demo

Copyright 2018 Giuseppe Pischedda
All rights reserved

https://www.software-on-demand-ita.com

** ** ** ** ** ** ** ** ** ** ** **/
#pragma once
#include "stdafx.h"


class CStaticImage : public CImage
{
public:
	
	CStaticImage();
	
	CStaticImage(LPCTSTR fileName);
	
	CStaticImage(HBITMAP hBm);
	
	virtual ~CStaticImage();

	HBITMAP GetSafeHBITMAP();
		
	CStaticImage(const CStaticImage &img);

	int Width();
	int Height();

		
	typedef struct tagSCALE_PERCENT
	{
		int cWidth;
		int cHeight;
	}SCALEPERCENT,*LPSCALEPERCENT;

	
	BOOL ScaleImage(int percent, SCALEPERCENT & scale);

	CStaticImage operator=(CStaticImage &obj);

	HBITMAP GetImageCopy();

	//Make a CImage copy
	HBITMAP  Clone();
	
	CImage Image;
	
	CImage GetClonedImage();
	
	HBITMAP GetClonedSafeHBITMAP();
	
	VOID SetImage(HBITMAP hBm);

	



private:
	CImage m_Image;
	CStaticImage *pImage;

	
	
};