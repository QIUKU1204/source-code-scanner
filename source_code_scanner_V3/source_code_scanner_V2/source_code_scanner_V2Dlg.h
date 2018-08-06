
// source_code_scanner_V2Dlg.h : ͷ�ļ�
//

#pragma once
#include "afxwin.h"


// SrcCodeScannerDlg �Ի���
class SrcCodeScannerDlg : public CDialogEx
{
// ����
public:
	SrcCodeScannerDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_SOURCE_CODE_SCANNER_V2_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
	afx_msg void OnBnClickedButtonSelectFile();
	afx_msg void OnBnClickedButtonSelectFolder();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	void OnTimer(UINT_PTR nIDEvent); 
	CEdit m_edit_filepath;
};
