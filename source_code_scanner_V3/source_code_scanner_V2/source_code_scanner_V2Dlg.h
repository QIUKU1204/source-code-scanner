
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
	afx_msg void OnBnClickedButtonMd();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnEnChangeEditTop();
	CEdit m_edit_filepath;
	afx_msg void OnBnClickedButtonSelectFile();
	afx_msg void OnBnClickedButtonSelectFolder();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnEnChangeEditHeader();
	CEdit m_edit_header;
	afx_msg void OnEnUpdateEditFooter();
	CEdit m_edit_footer;
	afx_msg void OnBnClickedCheckH();
	afx_msg void OnBnClickedCheckHpp();
	afx_msg void OnBnClickedCheckHxx();
	afx_msg void OnBnClickedCheckC();
	afx_msg void OnBnClickedCheckCpp();
	afx_msg void OnBnClickedCheckNone();
	CButton m_check_h;
	CButton m_check_hpp;
	CButton m_check_hxx;
	CButton m_check_c;
	CButton m_check_cpp;
	CButton m_check_none;
};
