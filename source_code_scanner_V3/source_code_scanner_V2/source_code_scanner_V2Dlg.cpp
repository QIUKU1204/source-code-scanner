
// source_code_scanner_V2Dlg.cpp : 实现文件
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


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// SrcCodeScannerDlg 对话框

// 标准构造函数
SrcCodeScannerDlg::SrcCodeScannerDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(SrcCodeScannerDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON_COLORS);
	// 编码初始化为GBK
	encoding = "GBK"; 
}

// 显式析构函数
SrcCodeScannerDlg::~SrcCodeScannerDlg(){
	// 主动释放vector 占用的内存
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


// SrcCodeScannerDlg 消息处理程序

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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void SrcCodeScannerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR SrcCodeScannerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 初始化对话框
BOOL SrcCodeScannerDlg::OnInitDialog() 
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标


	// TODO: 在此添加额外的初始化代码
	((CButton *)GetDlgItem(IDC_RADIO_GBK))->SetCheck(TRUE); // 初始默认选中GBK

	SetTimer(1, 1000, NULL); // 设置定时器
	
	wordOpt.CreateApp(); // 在打开程序窗口的同时，启动Word程序


	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

/////////////////用户自主添加控件的消息处理函数///////////////////

void SrcCodeScannerDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	if (nIDEvent == 1)
	{
		CWnd * hWnd = FindWindow(NULL,_T("代码扫描器"));
		if (hWnd)
		{
			Sleep(1000); // 先阻塞1S再关闭
			hWnd->PostMessageW(WM_CLOSE,NULL,NULL);
		}
	}
	//EndDialog(IDOK);
	CDialogEx::OnTimer(nIDEvent);
}

///////////定义全局变量(封装为Dlg类的数据成员)//////////

////////////////////////////////////////////////////////
void SrcCodeScannerDlg::OnBnClickedButtonWord()
{
	// TODO: 在此添加控件通知处理程序代码

	HWND hWnd = AfxGetMainWnd()->m_hWnd; // 获取当前窗口的句柄

	if (!scanner.CheckPathVector(path_vc,hWnd,file_extensions)) // 校验容器中的每一个路径元素
	{
		return; // 若容器为空，则不再执行扫描操作
	}
	
	// 根据传入的头文件路径，在其目录下生成相应的Word文档
	MessageBox(_T("扫描开始！"),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
	for (size_t i = 0;i < path_vc.size();i++)
	{
		if (header_text != "" && footer_text != "") // 若不为空，则传入页眉页脚
		{
			scanner.GenerateWordDoc(path_vc[i],file_extensions,encoding,wordOpt,header_text,footer_text);
		} 
		else // 若为空，则使用默认参数
		{
			scanner.GenerateWordDoc(path_vc[i],file_extensions,encoding,wordOpt);
		}
	}
	MessageBox(_T("扫描完成！"),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
}

void SrcCodeScannerDlg::OnBnClickedButtonMd()
{
	// TODO: 在此添加控件通知处理程序代码

	HWND hWnd = AfxGetMainWnd()->m_hWnd; // 获取当前窗口的句柄

	if (!scanner.CheckPathVector(path_vc,hWnd,file_extensions)) // 校验容器中的每一个路径元素
	{
		return; // 若容器为空，则不再执行扫描操作
	}

	MessageBox(_T("扫描开始！"),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
	// 根据传入的头文件路径，在其目录下生成相应的Markdown文档
	USES_CONVERSION;
	for (size_t i = 0;i < path_vc.size();i++)
	{
		if (header_text != "" && footer_text != "") // 若不为空，则传入页眉页脚
		{
			scanner.GenerateMarkdownFile(path_vc[i],file_extensions,encoding,W2A(header_text),W2A(footer_text));
		} 
		else // 若为空，则使用默认参数
		{
			scanner.GenerateMarkdownFile(path_vc[i],file_extensions,encoding);
		}
	}
	MessageBox(_T("扫描完成！"),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
}

void SrcCodeScannerDlg::OnBnClickedButtonSelectFile()
{
	// TODO: 在此添加控件通知处理程序代码

	// 设置过滤器 ("name1|value1|name2|value2||")
	TCHAR type_Filter[] = _T("所有文件(*.*)|*.*|头文件(*.h)|*.h|*.txt|*.txt|*.doc|*.doc|*.cpp|*.cpp||");   

	// 构造打开文件对话框(是否打开(否则为保存),.....)   
	CFileDialog file_Dlg(TRUE, NULL, NULL, OFN_OVERWRITEPROMPT | OFN_ALLOWMULTISELECT, type_Filter, this);

	CString cstr_filepath; 
	CString multi_filepath;

	// 显示打开文件对话框   
	if (IDOK == file_Dlg.DoModal()) // 若选择了文件，才清空vector容器 
	{   
		// 清空vector容器，防止受到上一次拖拽/选择的影响
		path_vc.clear();

		POSITION file_name_pos = file_Dlg.GetStartPosition();
		while (file_name_pos != NULL) // 选择多个文件
		{
			cstr_filepath = file_Dlg.GetNextPathName(file_name_pos);
			multi_filepath = multi_filepath + cstr_filepath + "; ";
			// 将文件路径赋给一个全局变量，在OnBnClickedButtonWord()中使用
			USES_CONVERSION;
			path_vc.push_back(W2A(cstr_filepath)); // CStringW -> CSringA -> string
		}
		SetDlgItemTextW(IDC_EDIT_TOP, multi_filepath); // 在编辑框中显示多个文件的路径
	}   
}

void SrcCodeScannerDlg::OnBnClickedButtonSelectFolder()
{
	// TODO: 在此添加控件通知处理程序代码
	// wchar_t <=> WCHAR <=> TCHAR
	WCHAR * wchar_folderpath = new WCHAR[MAX_PATH];
	CString cstr_folderpath;
	
	// 定义、设置浏览信息对象
	BROWSEINFO browse_info;
	::ZeroMemory(&browse_info,sizeof(BROWSEINFO));
	browse_info.pidlRoot = 0;                                 // 要显示的文件夹对话框的根(Root)
	browse_info.lpszTitle = _T("请选择需要扫描的工程文件夹"); // 显示位于对话框左上角的标题
	browse_info.ulFlags = BIF_RETURNONLYFSDIRS |              // 指定对话框外观和功能的标志
						BIF_EDITBOX | BIF_DONTGOBELOWDOMAIN;
	browse_info.lpfn = NULL;                                  // 处理事件的回调函数

	// 显示浏览文件夹的对话框
	LPITEMIDLIST lpid_browse = ::SHBrowseForFolder(&browse_info);
	if (lpid_browse != NULL)
	{
		// 清空vector容器，防止受到上一次拖拽/选择的影响
		path_vc.clear();

		// 取得文件夹名
		if (::SHGetPathFromIDList(lpid_browse,wchar_folderpath))
		{
			cstr_folderpath = wchar_folderpath;
			USES_CONVERSION;
			path_vc.push_back(W2A(cstr_folderpath));
			SetDlgItemTextW(IDC_EDIT_TOP,cstr_folderpath);
		}
	}

	// 释放内存
	if (lpid_browse != NULL)
	{
		::CoTaskMemFree(lpid_browse);
	}
	delete [] wchar_folderpath;
}

void SrcCodeScannerDlg::OnDropFiles(HDROP hDropInfo)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值

	// 清空vector容器，防止受到上一次拖拽/选择的影响
	path_vc.clear();

	// wchar_t <=> WCHAR <=> TCHAR
	WCHAR * wchar_filepath = new WCHAR[MAX_PATH];
	CString multi_filepath;

	// DragQueryFile(拖拽文件信息，查询文件的索引号，文件路径缓冲区，缓冲区大小)
	// 当索引号参数为0xFFFFFFFF时，该函数的返回值为拖拽文件的数量
	int file_count = DragQueryFile(hDropInfo,0xFFFFFFFF,NULL,0); // 获取文件/文件夹的数量

	// 根据索引号依次从拖拽信息中获取文件路径，放入到缓冲区中
	for (int index = 0;index < file_count;++index)
	{
		DragQueryFile(hDropInfo,index,wchar_filepath,MAX_PATH);

		// 将WCHAR*转换为string后放入容器中
		USES_CONVERSION;
		path_vc.push_back(W2A(wchar_filepath)); // WCHAR* -> char* -> string

		// 获取多个文件的路径
		multi_filepath = multi_filepath + wchar_filepath + "; ";
	}
	// 使用Edit控件的成员变量m_edit_filepath，向编辑框发送文件路径信息
	//m_edit_filepath.SetWindowTextW(multi_filepath);
	SetDlgItemTextW(IDC_EDIT_TOP,multi_filepath);

	DragFinish(hDropInfo);
	delete [] wchar_filepath;
}

void SrcCodeScannerDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}

void SrcCodeScannerDlg::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码

	if (MessageBox(_T("确定要关闭程序吗？"),_T("提示信息"),MB_OK|MB_ICONINFORMATION|MB_YESNO) == IDNO)
	{
		return;
	} 
	else{   // == IDYES
		try // 用于解决"RPC服务器不可用"的BUG
		{
			wordOpt.AppClose();  // 在关闭程序窗口的同时，关闭Word程序
		}
		catch (COleException* e) // 捕获到COleException 时，执行catch 代码块
		{
			e->Delete(); // 删除Exception 对象，释放其拥有的堆内存，防止内存泄露
			MessageBox(_T("程序正在关闭..."),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
			CDialogEx::OnCancel();
		}
		// 没有捕获到COleException 时
		MessageBox(_T("程序正在关闭..."),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
		CDialogEx::OnCancel();
	}
}

void SrcCodeScannerDlg::OnEnChangeEditTop()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

void SrcCodeScannerDlg::OnEnChangeEditHeader() // 向编辑框中每输入一个字符，Change即响应一次
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	CString cstr_edit;
	m_edit_header.GetWindowTextW(cstr_edit);
	header_text = cstr_edit;
}

void SrcCodeScannerDlg::OnEnUpdateEditFooter() // 向编辑框中每输入一个字符，Update即响应一次
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数，以将 EM_SETEVENTMASK 消息发送到该控件，
	// 同时将 ENM_UPDATE 标志“或”运算到 lParam 掩码中。

	// TODO:  在此添加控件通知处理程序代码
	CString cstr_edit;
	m_edit_footer.GetWindowTextW(cstr_edit);
	footer_text = cstr_edit;
}

void SrcCodeScannerDlg::OnBnClickedCheckH()
{
	// TODO: 在此添加控件通知处理程序代码
	int state = m_check_h.GetCheck();
	if (state == 1) // 当前复选框被选中
	{
		file_extensions.push_back(".h");
	}
	else{           // 若没有被选中
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
	// TODO: 在此添加控件通知处理程序代码
	int state = m_check_hpp.GetCheck();
	if (state == 1) // 当前复选框被选中
	{
		file_extensions.push_back(".hpp");
	}
	else{           // 若没有被选中
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
	// TODO: 在此添加控件通知处理程序代码
	int state = m_check_hxx.GetCheck();
	if (state == 1) // 当前复选框被选中
	{
		file_extensions.push_back(".hxx");
	}
	else{           // 若没有被选中
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
	// TODO: 在此添加控件通知处理程序代码
	int state = m_check_c.GetCheck();
	if (state == 1) // 当前复选框被选中
	{
		file_extensions.push_back(".c");
	}
	else{           // 若没有被选中
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
	// TODO: 在此添加控件通知处理程序代码
	int state = m_check_cpp.GetCheck();
	if (state == 1) // 当前复选框被选中
	{
		file_extensions.push_back(".cpp");
	}
	else{           // 若没有被选中
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
	// TODO: 在此添加控件通知处理程序代码
	//UpdateData(TRUE);
	int state = m_check_none.GetCheck();
	if (state == 1) // 当前复选框被选中
	{
		file_extensions.push_back("");
	}
	else{           // 若没有被选中
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
	// TODO: 在此添加控件通知处理程序代码
	encoding = "UTF-8";
}


void SrcCodeScannerDlg::OnBnClickedRadioGbk()
{
	// TODO: 在此添加控件通知处理程序代码
	encoding = "GBK";
}
