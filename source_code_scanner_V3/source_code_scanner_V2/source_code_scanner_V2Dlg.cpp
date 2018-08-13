
// source_code_scanner_V2Dlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "source_code_scanner_V2.h"
#include "source_code_scanner_V2Dlg.h"
//#include "afxdialogex.h"
#include "Word_Api.h"
#include "src_code_scanner.h"
#include <string>
#include <algorithm>
#include <exception>
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
	// �����ʼ��ΪGBK
	encoding = "GBK"; 
}

// ��ʽ��������
SrcCodeScannerDlg::~SrcCodeScannerDlg(){
	// �����ͷ�vector ռ�õ��ڴ�
	vector<string>().swap(file_extensions);
	vector<string>().swap(path_vc);
}

void SrcCodeScannerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_TOP, m_edit_filepath);
	DDX_Control(pDX, IDC_EDIT_HEADER, m_edit_header);
	DDX_Control(pDX, IDC_EDIT_FOOTER, m_edit_footer);
	DDX_Control(pDX, IDC_CHECK_H, m_check_h);
	DDX_Control(pDX, IDC_CHECK_HPP, m_check_hpp);
	DDX_Control(pDX, IDC_CHECK_HXX, m_check_hxx);
	DDX_Control(pDX, IDC_CHECK_C, m_check_c);
	DDX_Control(pDX, IDC_CHECK_CPP, m_check_cpp);
	DDX_Control(pDX, IDC_CHECK_NONE, m_check_none);
	DDX_Control(pDX, IDCANCEL, m_radio_encoding);
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
	ON_WM_TIMER()
	ON_EN_CHANGE(IDC_EDIT_HEADER, &SrcCodeScannerDlg::OnEnChangeEditHeader)
	ON_EN_UPDATE(IDC_EDIT_FOOTER, &SrcCodeScannerDlg::OnEnUpdateEditFooter)
	ON_BN_CLICKED(IDC_CHECK_H, &SrcCodeScannerDlg::OnBnClickedCheckH)
	ON_BN_CLICKED(IDC_CHECK_HPP, &SrcCodeScannerDlg::OnBnClickedCheckHpp)
	ON_BN_CLICKED(IDC_CHECK_HXX, &SrcCodeScannerDlg::OnBnClickedCheckHxx)
	ON_BN_CLICKED(IDC_CHECK_C, &SrcCodeScannerDlg::OnBnClickedCheckC)
	ON_BN_CLICKED(IDC_CHECK_CPP, &SrcCodeScannerDlg::OnBnClickedCheckCpp)
	ON_BN_CLICKED(IDC_CHECK_NONE, &SrcCodeScannerDlg::OnBnClickedCheckNone)
	ON_BN_CLICKED(IDC_RADIO_UTF8, &SrcCodeScannerDlg::OnBnClickedRadioUtf8)
	ON_BN_CLICKED(IDC_RADIO_GBK, &SrcCodeScannerDlg::OnBnClickedRadioGbk)
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

// ��ʼ���Ի���
BOOL SrcCodeScannerDlg::OnInitDialog() 
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
	((CButton *)GetDlgItem(IDC_RADIO_GBK))->SetCheck(TRUE); // ��ʼĬ��ѡ��GBK

	SetTimer(1, 1000, NULL); // ���ö�ʱ��
	
	wordOpt.CreateApp(); // �ڴ򿪳��򴰿ڵ�ͬʱ������Word����


	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

/////////////////�û�������ӿؼ�����Ϣ������///////////////////

void SrcCodeScannerDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	if (nIDEvent == 1)
	{
		CWnd * hWnd = FindWindow(NULL,_T("����ɨ����"));
		if (hWnd)
		{
			Sleep(1000); // ������1S�ٹر�
			hWnd->PostMessageW(WM_CLOSE,NULL,NULL);
		}
	}
	//EndDialog(IDOK);
	CDialogEx::OnTimer(nIDEvent);
}

///////////����ȫ�ֱ���(��װΪDlg������ݳ�Ա)//////////

////////////////////////////////////////////////////////
void SrcCodeScannerDlg::OnBnClickedButtonWord()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	HWND hWnd = AfxGetMainWnd()->m_hWnd; // ��ȡ��ǰ���ڵľ��

	if (!scanner.CheckPathVector(path_vc,hWnd,file_extensions)) // У�������е�ÿһ��·��Ԫ��
	{
		return; // ������Ϊ�գ�����ִ��ɨ�����
	}
	
	// ���ݴ����ͷ�ļ�·��������Ŀ¼��������Ӧ��Word�ĵ�
	MessageBox(_T("ɨ�迪ʼ��"),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
	for (size_t i = 0;i < path_vc.size();i++)
	{
		if (header_text != "" && footer_text != "") // ����Ϊ�գ�����ҳüҳ��
		{
			scanner.GenerateWordDoc(path_vc[i],file_extensions,encoding,wordOpt,header_text,footer_text);
		} 
		else // ��Ϊ�գ���ʹ��Ĭ�ϲ���
		{
			scanner.GenerateWordDoc(path_vc[i],file_extensions,encoding,wordOpt);
		}
	}
	MessageBox(_T("ɨ����ɣ�"),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
}

void SrcCodeScannerDlg::OnBnClickedButtonMd()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	HWND hWnd = AfxGetMainWnd()->m_hWnd; // ��ȡ��ǰ���ڵľ��

	if (!scanner.CheckPathVector(path_vc,hWnd,file_extensions)) // У�������е�ÿһ��·��Ԫ��
	{
		return; // ������Ϊ�գ�����ִ��ɨ�����
	}

	MessageBox(_T("ɨ�迪ʼ��"),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
	// ���ݴ����ͷ�ļ�·��������Ŀ¼��������Ӧ��Markdown�ĵ�
	USES_CONVERSION;
	for (size_t i = 0;i < path_vc.size();i++)
	{
		if (header_text != "" && footer_text != "") // ����Ϊ�գ�����ҳüҳ��
		{
			scanner.GenerateMarkdownFile(path_vc[i],file_extensions,encoding,W2A(header_text),W2A(footer_text));
		} 
		else // ��Ϊ�գ���ʹ��Ĭ�ϲ���
		{
			scanner.GenerateMarkdownFile(path_vc[i],file_extensions,encoding);
		}
	}
	MessageBox(_T("ɨ����ɣ�"),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
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
	WCHAR * wchar_folderpath = new WCHAR[MAX_PATH];
	CString cstr_folderpath;
	
	// ���塢���������Ϣ����
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
	WCHAR * wchar_filepath = new WCHAR[MAX_PATH];
	CString multi_filepath;

	// DragQueryFile(��ק�ļ���Ϣ����ѯ�ļ��������ţ��ļ�·������������������С)
	// �������Ų���Ϊ0xFFFFFFFFʱ���ú����ķ���ֵΪ��ק�ļ�������
	int file_count = DragQueryFile(hDropInfo,0xFFFFFFFF,NULL,0); // ��ȡ�ļ�/�ļ��е�����

	// �������������δ���ק��Ϣ�л�ȡ�ļ�·�������뵽��������
	for (int index = 0;index < file_count;++index)
	{
		DragQueryFile(hDropInfo,index,wchar_filepath,MAX_PATH);

		// ��WCHAR*ת��Ϊstring�����������
		USES_CONVERSION;
		path_vc.push_back(W2A(wchar_filepath)); // WCHAR* -> char* -> string

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
	CDialogEx::OnOK();
}

void SrcCodeScannerDlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	if (MessageBox(_T("ȷ��Ҫ�رճ�����"),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION|MB_YESNO) == IDNO)
	{
		return;
	} 
	else{   // == IDYES
		try // ���ڽ��"RPC������������"��BUG
		{
			wordOpt.AppClose();  // �ڹرճ��򴰿ڵ�ͬʱ���ر�Word����
		}
		catch (COleException* e) // ����COleException ʱ��ִ��catch �����
		{
			e->Delete(); // ɾ��Exception �����ͷ���ӵ�еĶ��ڴ棬��ֹ�ڴ�й¶
			MessageBox(_T("�������ڹر�..."),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
			CDialogEx::OnCancel();
		}
		// û�в���COleException ʱ
		MessageBox(_T("�������ڹر�..."),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
		CDialogEx::OnCancel();
	}
}

void SrcCodeScannerDlg::OnEnChangeEditTop()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

void SrcCodeScannerDlg::OnEnChangeEditHeader() // ��༭����ÿ����һ���ַ���Change����Ӧһ��
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString cstr_edit;
	m_edit_header.GetWindowTextW(cstr_edit);
	header_text = cstr_edit;
}

void SrcCodeScannerDlg::OnEnUpdateEditFooter() // ��༭����ÿ����һ���ַ���Update����Ӧһ��
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// �������Խ� EM_SETEVENTMASK ��Ϣ���͵��ÿؼ���
	// ͬʱ�� ENM_UPDATE ��־�������㵽 lParam �����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString cstr_edit;
	m_edit_footer.GetWindowTextW(cstr_edit);
	footer_text = cstr_edit;
}

void SrcCodeScannerDlg::OnBnClickedCheckH()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int state = m_check_h.GetCheck();
	if (state == 1) // ��ǰ��ѡ��ѡ��
	{
		file_extensions.push_back(".h");
	}
	else{           // ��û�б�ѡ��
		for (size_t i = 0;i < file_extensions.size();i++)
		{
			if (file_extensions[i] == ".h")
			{
				file_extensions.erase(file_extensions.begin() + i);
			}
		}
	}
}

void SrcCodeScannerDlg::OnBnClickedCheckHpp()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int state = m_check_hpp.GetCheck();
	if (state == 1) // ��ǰ��ѡ��ѡ��
	{
		file_extensions.push_back(".hpp");
	}
	else{           // ��û�б�ѡ��
		for (size_t i = 0;i < file_extensions.size();i++)
		{
			if (file_extensions[i] == ".hpp")
			{
				file_extensions.erase(file_extensions.begin() + i);
			}
		}
	}
}

void SrcCodeScannerDlg::OnBnClickedCheckHxx()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int state = m_check_hxx.GetCheck();
	if (state == 1) // ��ǰ��ѡ��ѡ��
	{
		file_extensions.push_back(".hxx");
	}
	else{           // ��û�б�ѡ��
		for (size_t i = 0;i < file_extensions.size();i++)
		{
			if (file_extensions[i] == ".hxx")
			{
				file_extensions.erase(file_extensions.begin() + i);
			}
		}
	}
}

void SrcCodeScannerDlg::OnBnClickedCheckC()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int state = m_check_c.GetCheck();
	if (state == 1) // ��ǰ��ѡ��ѡ��
	{
		file_extensions.push_back(".c");
	}
	else{           // ��û�б�ѡ��
		for (size_t i = 0;i < file_extensions.size();i++)
		{
			if (file_extensions[i] == ".c")
			{
				file_extensions.erase(file_extensions.begin() + i);
			}
		}
	}
}

void SrcCodeScannerDlg::OnBnClickedCheckCpp()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int state = m_check_cpp.GetCheck();
	if (state == 1) // ��ǰ��ѡ��ѡ��
	{
		file_extensions.push_back(".cpp");
	}
	else{           // ��û�б�ѡ��
		for (size_t i = 0;i < file_extensions.size();i++)
		{
			if (file_extensions[i] == ".cpp")
			{
				file_extensions.erase(file_extensions.begin() + i);
			}
		}
	}
}

void SrcCodeScannerDlg::OnBnClickedCheckNone()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	//UpdateData(TRUE);
	int state = m_check_none.GetCheck();
	if (state == 1) // ��ǰ��ѡ��ѡ��
	{
		file_extensions.push_back("");
	}
	else{           // ��û�б�ѡ��
		for (size_t i = 0;i < file_extensions.size();i++)
		{
			if (file_extensions[i] == "")
			{
				file_extensions.erase(file_extensions.begin() + i);
			}
		}
	}
}

void SrcCodeScannerDlg::OnBnClickedRadioUtf8()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	encoding = "UTF-8";
}


void SrcCodeScannerDlg::OnBnClickedRadioGbk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	encoding = "GBK";
}
