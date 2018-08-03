
// source_code_scanner_V2Dlg.h : 头文件
//

#pragma once
#include "afxwin.h"


// SrcCodeScannerDlg 对话框
class SrcCodeScannerDlg : public CDialogEx
{
// 构造
public:
	SrcCodeScannerDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_SOURCE_CODE_SCANNER_V2_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonWord();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnEnChangeEdit1();
	afx_msg void OnBnClickedButtonSelect();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	CEdit m_edit_filepath;
};
