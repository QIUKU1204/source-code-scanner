
// source_code_scanner_V2Dlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "source_code_scanner_V2.h"
#include "source_code_scanner_V2Dlg.h"
#include "afxdialogex.h"
#include "Word_Api.h"
#include "src_code_scanner.h"
#include <string>
#include <algorithm>
using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// SrcCodeScannerDlg �Ի���

// ��׼���캯��
SrcCodeScannerDlg::SrcCodeScannerDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(SrcCodeScannerDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON_COLORS);
}

void SrcCodeScannerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_TOP, m_edit_filepath);
}

BEGIN_MESSAGE_MAP(SrcCodeScannerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_DROPFILES()
	ON_BN_CLICKED(IDOK, &SrcCodeScannerDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &SrcCodeScannerDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BUTTON_WORD, &SrcCodeScannerDlg::OnBnClickedButtonWord)
	ON_BN_CLICKED(IDC_BUTTON_MD, &SrcCodeScannerDlg::OnBnClickedButtonMd)
	ON_EN_CHANGE(IDC_EDIT_TOP, &SrcCodeScannerDlg::OnEnChangeEditTop)
	ON_BN_CLICKED(IDC_BUTTON_SELECT_FILE, &SrcCodeScannerDlg::OnBnClickedButtonSelectFile)
	ON_BN_CLICKED(IDC_BUTTON_SELECT_FOLDER, &SrcCodeScannerDlg::OnBnClickedButtonSelectFolder)
END_MESSAGE_MAP()


// SrcCodeScannerDlg ��Ϣ�������

void SrcCodeScannerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void SrcCodeScannerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR SrcCodeScannerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

SccWordApi wordOpt; // Word����������Ϊȫ��
BOOL SrcCodeScannerDlg::OnInitDialog() // ��ʼ���Ի���
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��


	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	SetTimer(1, 2000, NULL); // ���ö�ʱ��

	wordOpt.CreateApp(); // �ڴ򿪳��򴰿ڵ�ͬʱ������Word����


	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void SrcCodeScannerDlg::OnTimer(UINT_PTR nIDEvent){
	// Ŀǰ�޷��Զ��ر���Ϣ�Ի���
}

/////////////////�û�������ӿؼ�����Ϣ������///////////////////

vector<string> path_vc; // �ļ�&�ļ���·��������Ϊȫ��
void SrcCodeScannerDlg::OnBnClickedButtonWord()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	SrcCodeScanner scanner; // ����һ��ɨ���������Ե��÷�װ�õĺ���

	HWND hWnd = AfxGetMainWnd()->m_hWnd; // ��ȡ��ǰ���ڵľ��

	if (!scanner.CheckPathVector(path_vc,hWnd)) // ��������Ƿ�Ϊ��
	{
		return; // ��Ϊ�գ�����ִ��ɨ�����
	}
	
	MessageBox(_T("ɨ�迪ʼ��"),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);
	// ���ݴ����ͷ�ļ�·��������Ŀ¼��������Ӧ��Word�ĵ�
	//wordOpt.CreateApp(); // OnInitDialog()
	for (unsigned int i = 0;i < path_vc.size();i++)
	{
		scanner.GenerateWordDoc(path_vc[i],wordOpt);
	}
	MessageBox(_T("ɨ����ɣ�"),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);
	//wordOpt.AppClose(); // OnBnClickedOk(); OnBnClickedCancel()
}

void SrcCodeScannerDlg::OnBnClickedButtonMd()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	MessageBox(_T("MD�ļ����ܴ�����...�����ڴ���"),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);

	// TODO: �ڴ���ӿؼ�֪ͨ����������
	/*SrcCodeScanner scanner; // ����һ��ɨ���������Ե��÷�װ�õĺ���

	HWND hWnd = AfxGetMainWnd()->m_hWnd; // ��ȡ��ǰ���ڵľ��

	if (!scanner.CheckPathVector(path_vc,hWnd)) // ��������Ƿ�Ϊ��
	{
		return; // ��Ϊ�գ�����ִ��ɨ�����
	}

	MessageBox(_T("ɨ�迪ʼ��"),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);
	// ���ݴ����ͷ�ļ�·��������Ŀ¼��������Ӧ��Word�ĵ�
	for (unsigned int i = 0;i < path_vc.size();i++)
	{
		scanner.GenerateMarkdownFile(path_vc[i]);
	}
	MessageBox(_T("ɨ����ɣ�"),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);*/
	
}

void SrcCodeScannerDlg::OnBnClickedButtonSelectFile()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	// ���ù����� ("name1|value1|name2|value2||")
	TCHAR type_Filter[] = _T("�����ļ�(*.*)|*.*|ͷ�ļ�(*.h)|*.h|*.txt|*.txt|*.doc|*.doc|*.cpp|*.cpp||");   

	// ������ļ��Ի���(�Ƿ��(����Ϊ����),.....)   
	CFileDialog file_Dlg(TRUE, NULL, NULL, OFN_OVERWRITEPROMPT | OFN_ALLOWMULTISELECT, type_Filter, this);

	CString cstr_filepath; 
	CString multi_filepath;

	// ��ʾ���ļ��Ի���   
	if (IDOK == file_Dlg.DoModal()) // ��ѡ�����ļ��������vector���� 
	{   
		// ���vector��������ֹ�ܵ���һ����ק/ѡ���Ӱ��
		path_vc.clear();

		POSITION file_name_pos = file_Dlg.GetStartPosition();
		while (file_name_pos != NULL) // ѡ�����ļ�
		{
			cstr_filepath = file_Dlg.GetNextPathName(file_name_pos);
			multi_filepath = multi_filepath + cstr_filepath + "; ";
			// ���ļ�·������һ��ȫ�ֱ�������OnBnClickedButtonWord()��ʹ��
			USES_CONVERSION;
			path_vc.push_back(W2A(cstr_filepath)); // CStringW -> CSringA -> string
		}
		SetDlgItemTextW(IDC_EDIT_TOP, multi_filepath); // �ڱ༭������ʾ����ļ���·��
	}   
}

void SrcCodeScannerDlg::OnBnClickedButtonSelectFolder()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	// wchar_t <=> WCHAR <=> TCHAR
	WCHAR *wchar_folderpath = new WCHAR[MAX_PATH];
	CString cstr_folderpath;
	
	// ���������Ϣ����
	BROWSEINFO browse_info;
	::ZeroMemory(&browse_info,sizeof(BROWSEINFO));
	browse_info.pidlRoot = 0;                                 // Ҫ��ʾ���ļ��жԻ���ĸ�(Root)
	browse_info.lpszTitle = _T("��ѡ����Ҫɨ��Ĺ����ļ���"); // ��ʾλ�ڶԻ������Ͻǵı���
	browse_info.ulFlags = BIF_RETURNONLYFSDIRS |              // ָ���Ի�����ۺ͹��ܵı�־
						BIF_EDITBOX | BIF_DONTGOBELOWDOMAIN;
	browse_info.lpfn = NULL;                                  // �����¼��Ļص�����

	// ��ʾ����ļ��еĶԻ���
	LPITEMIDLIST lpid_browse = ::SHBrowseForFolder(&browse_info);
	if (lpid_browse != NULL)
	{
		// ���vector��������ֹ�ܵ���һ����ק/ѡ���Ӱ��
		path_vc.clear();

		// ȡ���ļ�����
		if (::SHGetPathFromIDList(lpid_browse,wchar_folderpath))
		{
			cstr_folderpath = wchar_folderpath;
			USES_CONVERSION;
			path_vc.push_back(W2A(cstr_folderpath));
			SetDlgItemTextW(IDC_EDIT_TOP,cstr_folderpath);
		}
	}

	// �ͷ��ڴ�
	if (lpid_browse != NULL)
	{
		::CoTaskMemFree(lpid_browse);
	}
	delete [] wchar_folderpath;
}

void SrcCodeScannerDlg::OnDropFiles(HDROP hDropInfo)
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ

	// ���vector��������ֹ�ܵ���һ����ק/ѡ���Ӱ��
	path_vc.clear();

	// wchar_t <=> WCHAR <=> TCHAR
	WCHAR *wchar_filepath = new WCHAR[MAX_PATH];
	CString multi_filepath;

	// DragQueryFile(��ק�ļ���Ϣ����ѯ�ļ��������ţ��ļ�·������������������С)
	// �������Ų���Ϊ0xFFFFFFFFʱ���ú����ķ���ֵΪ��ק�ļ�������
	int file_count = DragQueryFile(hDropInfo,0xFFFFFFFF,NULL,0); // ��ȡ�ļ�/�ļ��е�����

	// �������������δ���ק��Ϣ�л�ȡ�ļ�·�������뵽��������
	for (int index = 0;index < file_count;++index)
	{
		DragQueryFile(hDropInfo,index,wchar_filepath,MAX_PATH);

		// ��WCHARת��Ϊstring�����������
		string str_tmp;
		SrcCodeScanner scanner;
		scanner.WCHARToString(wchar_filepath,str_tmp);
		path_vc.push_back(str_tmp);

		// ��ȡ����ļ���·��
		multi_filepath = multi_filepath + wchar_filepath + "; ";
	}
	// ʹ��Edit�ؼ��ĳ�Ա����m_edit_filepath����༭�����ļ�·����Ϣ
	//m_edit_filepath.SetWindowTextW(multi_filepath);
	SetDlgItemTextW(IDC_EDIT_TOP,multi_filepath);

	DragFinish(hDropInfo);
	delete [] wchar_filepath;
}

void SrcCodeScannerDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	wordOpt.AppClose(); // �ڹرճ��򴰿ڵ�ͬʱ���ر�Word����
	CDialogEx::OnOK();
}

void SrcCodeScannerDlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	wordOpt.AppClose(); // �ڹرճ��򴰿ڵ�ͬʱ���ر�Word����
	CDialogEx::OnCancel();
}

void SrcCodeScannerDlg::OnEnChangeEditTop()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

// ��༭����ÿ����һ���ַ���Update������Ӧһ��
/*void SrcCodeScannerDlg::OnEnUpdateEditTop()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// �������Խ� EM_SETEVENTMASK ��Ϣ���͵��ÿؼ���
	// ͬʱ�� ENM_UPDATE ��־�������㵽 lParam �����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������

	// ���vector��������ֹ�ܵ���һ����ק/ѡ���Ӱ��
	filenames_vc.clear(); // ���޴���䣬Word�ĵ����ɽ�����ܳ���ѭ��

	CString cstr_edit;
	//UpdateData(TRUE);
	m_edit_filepath.GetWindowTextW(cstr_edit);
	USES_CONVERSION;
	//string str_tmp(W2A(cstr_edit)); // CStringW -> CSringA -> string
	filenames_vc.push_back(W2A(cstr_edit));
}*/
