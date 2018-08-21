
// source_code_scanner_V2Dlg.h : ͷ�ļ�
//

#pragma once
#include "afxwin.h"
#include "Word_Api.h"
#include "src_code_scanner.h"

// SrcCodeScannerDlg �Ի���
class SrcCodeScannerDlg : public CDialogEx
{
// ����
public:
	SrcCodeScannerDlg(CWnd* pParent = NULL);	// ��׼���캯��
	~SrcCodeScannerDlg();                       // ��ʽ��������

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
	// ��afx_msg �ؼ��ִ�ͷ����Ϣ������
	afx_msg void OnBnClickedButtonWord();
	afx_msg void OnBnClickedButtonMd();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnEnChangeEditTop();
	CString m_edit_filepath;
	afx_msg void OnBnClickedButtonSelectFile();
	afx_msg void OnBnClickedButtonSelectFolder();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnEnChangeEditHeader();
	CString m_edit_header;
	afx_msg void OnEnUpdateEditFooter();
	CString m_edit_footer;
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
	afx_msg void OnBnClickedRadioUtf8();
	afx_msg void OnBnClickedRadioGbk();
	CButton m_radio_encoding;
	afx_msg void OnBnClickedButtonInstructions();

private:
	SrcCodeScanner scanner;         // ɨ��������
	SccWordApi wordOpt;             // Word��������
	string encoding;                // �����ĵ�����
	vector<string> file_extensions; // �ļ���׺/��չ��
	vector<string> path_vc;         // �ļ�&�ļ���·��
};
