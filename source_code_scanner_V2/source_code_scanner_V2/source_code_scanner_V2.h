
// source_code_scanner_V2.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// SrcCodeScannerApp:
// �йش����ʵ�֣������ source_code_scanner_V2.cpp
//

class SrcCodeScannerApp : public CWinApp
{
public:
	SrcCodeScannerApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern SrcCodeScannerApp theApp;