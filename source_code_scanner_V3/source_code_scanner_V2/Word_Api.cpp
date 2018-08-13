#include "stdafx.h"
#include "Word_Api.h"
#include "src_code_scanner.h"


_Application    m_app;           ///< 创建word
Documents       m_docs;          ///< word文档集合
_Document       m_doc;           ///< 一个word文件
_Font           m_font;          ///< 字体对象
Selection       m_sel;           ///< 选择编辑对象，没有对象的时候就是插入点
Tables          m_tabs;
Table           m_tab;           ///< 表格对象
Bookmarks       m_bmarks;        
Bookmark        m_bmark;         ///< 标签对象
Hyperlinks      m_hlinks;
Hyperlink       m_hlink;         ///< 超链接对象
InlineShapes    m_ishapes;       
InlineShape     m_ishape;        ///< 图片对象
Borders         m_bds;
Border          m_bd;            ///< 表格边框对象
Window			m_win;           ///< 窗口对象
Pane			m_pane;			 ///< Pane(窗格)对象
View			m_view;			 ///< 视图对象
HeaderFooter	m_hefo;			 ///< 页眉页脚对象
List			m_list;          ///< List对象
Field           m_field;         
Fields			m_fields;		 ///< 域对象


SccWordApi::SccWordApi(void){}


SccWordApi::~SccWordApi(void){

	CoUninitialize();

	m_app.ReleaseDispatch();
	m_doc.ReleaseDispatch();
	m_docs.ReleaseDispatch();
	m_font.ReleaseDispatch();
	m_sel.ReleaseDispatch();
	m_tabs.ReleaseDispatch();
	m_tab.ReleaseDispatch();
	m_bmarks.ReleaseDispatch();
	m_bmark.ReleaseDispatch();
	m_hlink.ReleaseDispatch();
	m_hlinks.ReleaseDispatch();
	m_ishape.ReleaseDispatch();
	m_ishapes.ReleaseDispatch();
	m_win.ReleaseDispatch();
	m_pane.ReleaseDispatch();
	m_view.ReleaseDispatch();
	m_hefo.ReleaseDispatch();
	m_list.ReleaseDispatch();
	m_bd.ReleaseDispatch();
	m_bds.ReleaseDispatch();
	m_field.ReleaseDispatch();
	m_fields.ReleaseDispatch();

}


BOOL SccWordApi::InitCOM(){
	// 初始化COM
	if (!AfxOleInit())
	{
		AfxMessageBox(_T("无法初始化COM的动态链接库"));
		return FALSE;
	}
	else{
		AfxMessageBox(_T("COM初始化成功！"));
		return TRUE;
	}
}

BOOL SccWordApi::CreateApp(){
	if (!m_app.CreateDispatch(_T("Word.Application")))
	{
		AfxMessageBox(_T("服务创建失败，请确定已经安装了Office 2000或以上版本"));
		return FALSE;
	} 
	else
	{
		//m_app.put_Visible(TRUE);

		//AfxMessageBox(_T("Word启动成功，按确认键继续"),MB_OK|MB_ICONINFORMATION);

		return TRUE;
	}
}

BOOL SccWordApi::CreateDocument(){
	if (!m_app.m_lpDispatch)
	{
		AfxMessageBox(_T("没有打开Word程序，文档创建失败！"), MB_OK | MB_ICONWARNING);
		return FALSE;
	} 
	else
	{
		try // 用于解决"RPC服务器不可用"的BUG
		{
			m_docs = m_app.get_Documents();
		}
		catch (COleException* e)
		{
			throw e; // 将异常往上抛出，由GenerateWordDoc() 捕获并处理
		}
		
		if (m_docs.m_lpDispatch == NULL)
		{
			AfxMessageBox(_T("Word文档集创建失败！"));
			return FALSE;
		} 
		else
		{
			COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
			m_doc = m_docs.Add(vOpt,vOpt,vOpt,vOpt);

			if (!m_doc.m_lpDispatch)
			{
				AfxMessageBox(_T("Word文档创建失败！"));
				return FALSE;
			} 
			else
			{
				m_sel = m_app.get_Selection(); // 获得当前Word操作
				if (!m_sel.m_lpDispatch)
				{
					AfxMessageBox(_T("选择编辑对象Selection获取失败！"));
					return FALSE;
				} 
				else
				{
					//AfxMessageBox(_T("Word文档创建成功！"));
					return TRUE;
				}
			}
		}
	}
}

BOOL SccWordApi::Create(){
	if (CreateApp())
	{
		if (CreateDocument())
		{
			return TRUE;
		} 
		else
		{
			return FALSE;
		}
	} 
	else
	{
		return FALSE;
	}
}

BOOL SccWordApi::Open(CString FileName,BOOL ReadOnly /* = FALSE */,BOOL AddToRecentFiles /* = FALSE */){
	CComVariant Read(ReadOnly),AddToR(AddToRecentFiles),Name(FileName);
	CComVariant format(0);
	COleVariant vTrue((short)TRUE),vFalse((short)FALSE);
	COleVariant varstrNull(_T(""));
	COleVariant varFalse(short(0),VT_BOOL);
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);

	if (!m_app.m_lpDispatch)
	{
		if(CreateApp()==FALSE)
			return FALSE;
	}
	if (!m_docs.m_lpDispatch)
	{
		m_docs = m_app.get_Documents();
		if (!m_docs.m_lpDispatch)
		{
			AfxMessageBox(_T("Word文档集创建失败！"));
			return FALSE;
		}
	}
	m_doc = m_docs.Open(&Name,varFalse,&Read,&AddToR,vOpt,vOpt,
		vFalse,vOpt,vOpt,&format,vOpt,vTrue,vOpt,vOpt,vOpt,vOpt);
	if (!m_doc.m_lpDispatch)
	{
		AfxMessageBox(_T("Word文档创建失败！"));
		return FALSE;
	} 
	else
	{
		m_sel = m_app.get_Selection();
		if (!m_sel.m_lpDispatch)
		{
			AfxMessageBox(_T("Word文档打开失败！"));
			return FALSE;
		}
		//AfxMessageBox(_T("Word文档打开成功！"));
		return TRUE;
	}
}

BOOL SccWordApi::Close(BOOL SaveChange/* =FALSE */)
{
	CComVariant vTrue(TRUE);
	CComVariant vFalse(FALSE);
	CComVariant vOpt;
	CComVariant cSavechage(SaveChange);
	if(!m_doc.m_lpDispatch)
	{
		AfxMessageBox(_T("Word文档关闭失败！"));
		return FALSE;
	}
	else
	{
		if(TRUE==SaveChange)
		{
			Save();
		}
		//下面第一个参数填vTrue 会出现错误，可能是后面的参数也要对应的变化
		//但vba 没有给对应参数 我就用这种方法来保存
		m_doc.Close(&vFalse,&vOpt,&vOpt);
		//AfxMessageBox(_T("Word文档关闭成功！"));
	}
	return TRUE;
}

BOOL SccWordApi::Save(){
	if (!m_doc.m_lpDispatch)
	{
		AfxMessageBox(_T("没有创建Document对象，保存失败"));
		return FALSE;
	} 
	else
	{
		m_doc.Save();
		//AfxMessageBox(_T("保存成功！"));
		return TRUE;
	}
}

BOOL SccWordApi::SaveAs(CString FileName,int SaveType /* = 0 */){
	CComVariant vTrue(TRUE);
	CComVariant vFalse(FALSE);
	CComVariant vOpt;
	CComVariant cFileName(FileName);
	CComVariant FileFormat(SaveType);
	m_doc = m_app.get_ActiveDocument();
	if (!m_doc.m_lpDispatch)
	{
		AfxMessageBox(_T("没有创建Document对象，保存失败"));
		return FALSE;
	} 
	else
	{
		// 最好按照宏来编写，不然可能出错
		m_doc.SaveAs(&cFileName,&FileFormat,&vFalse,COleVariant(_T("")),&vTrue,
			COleVariant(_T("")),&vFalse,&vFalse,&vFalse,&vFalse,&vFalse,&vOpt,&vOpt,&vOpt,&vOpt,&vOpt);
		//AfxMessageBox(_T("Word文档保存成功！"));
		return TRUE;
	}
}



void SccWordApi::WriteText(CString Text){
	m_sel.TypeText(Text);
}
void SccWordApi::NewLine(int nCount /* = 1 */){
	if (nCount <= 0)
	{
		nCount = 0;
	} 
	else
	{
		for (int i = 0;i < nCount;i++)
		{
			m_sel.TypeParagraph(); // 新建一段
		}
	}
}
void SccWordApi::WriteTextAfterNewLine(CString Text,int nCount /* = 1 */){
	NewLine(nCount);
	WriteText(Text);
}
void SccWordApi::EndLine(){
	m_sel = m_app.get_Selection();
	m_sel.EndKey(COleVariant((short)6),COleVariant((short)0)); // 将光标移动到最后一行
}
void SccWordApi::PageBreak(){
	m_sel = m_app.get_Selection();
	m_sel.InsertBreak(COleVariant((long)7)); // 插入分页符
}



void SccWordApi::SetFont(BOOL Bold,BOOL Italic /* = FALSE */,BOOL UnderLine /* = FALSE */){
	if (!m_sel.m_lpDispatch)
	{
		AfxMessageBox(_T("选择编辑对象为空，无法设置字体"));
		return;
	} 
	else
	{
		// const char *类型的实参与LPCTSTR类型的行参不兼容
		// m_sel.put_Text("F");
		m_sel.put_Text(_T("F")); 
		m_font = m_sel.get_Font(); // 获得字体对象
		m_font.put_Bold(Bold);  
		m_font.put_Italic(Italic);
		m_font.put_Underline(UnderLine); // 设置字体的各项参数
		m_sel.put_Font(m_font); //保存字体设置并生效
	}
}
void SccWordApi::SetFont(CString FontName,int FontSize /* = 9 */,long FontColor /* = 0 */,long Bold /* = 0 */){
	if (!m_sel.m_lpDispatch)
	{
		AfxMessageBox(_T("选择编辑对象为空，无法设置字体"));
		return;
	}
	// 使用 get_Font 获取字体对象
	// 使用 put_Text 获取字体属性
	m_sel.put_Text(_T("a"));
	m_font = m_sel.get_Font(); // 获取字体对象
	m_font.put_Size((float)FontSize);
	m_font.put_Name(FontName);
	m_font.put_Color(FontColor);
	m_font.put_Bold(Bold);
	m_sel.put_Font(m_font); // 选定字体对象
}
void SccWordApi::SetTableFont(int Row,int Column,CString FontName,int FontSize /* = 9 */,long FontColor /* = 0 */,long FontBackColor /* = 0 */){
	Cell cell = m_tab.Cell(Row,Column);
	cell.Select();
	_Font ft = m_sel.get_Font();
	ft.put_Name(FontName);
	ft.put_Size((float)FontSize);
	ft.put_Color(FontColor);
	Range ra = m_sel.get_Range();
	ra.put_HighlightColorIndex(FontBackColor);
}



void SccWordApi::CreateTable(int Row,int Column){
	m_doc = m_app.get_ActiveDocument();
	m_tabs = m_doc.get_Tables();
	CComVariant DefaultBehavior(1),AutoFitBehavior(2);

	if (!m_tabs.m_lpDispatch)
	{
		AfxMessageBox(_T("创建表格对象失败"));
		return;
	} 
	else
	{
		m_tabs.Add(m_sel.get_Range(),Row,Column,&DefaultBehavior,&AutoFitBehavior);
		m_tab = m_tabs.Item(m_tabs.get_Count()); // 

		// 设置表格的宽度(百分比%)
		m_tab.put_PreferredWidth(101.5);

		// 设置表格的边框属性(需要知道宏值)
		/*m_bds = m_tab.get_Borders();
		m_bds.put_OutsideLineStyle(20);
		m_bds.put_OutsideColorIndex(0);
		m_bds.put_InsideLineStyle(10);
		m_bds.put_InsideColorIndex(0);*/
	}
}
void SccWordApi::WriteCellText(int Row,int Column,CString Text){
	Cell cell = m_tab.Cell(Row,Column); // 从表格中获取指定单元格
	cell.Select(); // 选择该单元格作为编辑对象
	m_sel.TypeText(Text); // 将文本内容写入该单元格
}
void SccWordApi::SetTableShading(int Row,int Column,long ShadingColor /* = 0 */){
	Cell cell = m_tab.Cell(Row,Column);
	cell.Select();
	Shading	shading = cell.get_Shading(); // 从指定单元格获取其Shading(背景颜色)对象
	shading.put_BackgroundPatternColor(ShadingColor);
}



void SccWordApi::CreateBookMark(CString bmarkName)
{
	m_doc = m_app.get_ActiveDocument();
	m_sel = m_app.get_Selection();
	m_bmarks=m_doc.get_Bookmarks(); 

	if(!m_bmarks.m_lpDispatch)
	{
		AfxMessageBox(_T("创建标签对象失败"));
		return;
	}
	else
	{
		COleVariant vOpt;
		Range aRange = m_sel.get_Range();
		aRange.m_lpDispatch->AddRef();				///< 不加这句话会挂
		vOpt.vt = VT_DISPATCH;						//指定vt的类型
		vOpt.pdispVal = aRange;						//把range对象的m_lpDispatch赋值给vt
		m_bmark = m_bmarks.Add(bmarkName, &vOpt);	//插入一个叫bmarkName的书签
	}
}
void SccWordApi::CreateHyperLink(CString hlinkName,CString bmarkName)
{
	COleVariant     vAddress(_T("")),vSubAddress(bmarkName),vScreenTip(_T("")),vTextToDisplay(hlinkName);

	m_doc=m_app.get_ActiveDocument();
	Range aRange = m_sel.get_Range();

	m_hlinks = m_doc.get_Hyperlinks();
	m_hlinks.Add(
		aRange,				//Object，必需。转换为超链接的文本或图形。
		vAddress,			//Variant 类型，可选。指定的链接的地址。此地址可以是电子邮件地址、Internet 地址或文件名。请注意，Microsoft Word 不检查该地址的正确性。
		vSubAddress,		//Variant 类型，可选。目标文件内的位置名，如书签、已命名的区域或幻灯片编号。
		vScreenTip,			//Variant 类型，可选。当鼠标指针放在指定的超链接上时显示的可用作“屏幕提示”的文本。默认值为 Address。
		vTextToDisplay,		//Variant 类型，可选。指定的超链接的显示文本。此参数的值将取代由 Anchor 指定的文本或图形。
		vSubAddress			//Variant 类型，可选。要在其中打开指定的超链接的框架或窗口的名字。

		);
}
void SccWordApi::RepositionHyperLink(CString hlinkName,CString bmarkName)
{
	COleVariant     vSubAddress(bmarkName),vTextToDisplay(hlinkName);
	m_hlinks = m_doc.get_Hyperlinks();

	m_hlink = m_hlinks.Item(&vTextToDisplay);
	m_hlink.put_SubAddress(bmarkName);
}



void SccWordApi::InsertShapes(CString fileName)
{
	COleVariant vTrue((short)TRUE),

		vFalse((short)FALSE),

		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_ishapes=m_sel.get_InlineShapes();

	m_ishape =m_ishapes.AddPicture(fileName,vFalse,vTrue,vOpt);

}
void SccWordApi::InsertShapes(CBitmap *pBitmap)
{
	if (OpenClipboard(NULL))
	{
		EmptyClipboard();   
		SetClipboardData(CF_BITMAP,pBitmap->m_hObject);   
		CloseClipboard();   
		m_sel.Paste();
	}
}



void SccWordApi::SetParagraphFormat(int Alignment /* = 3 */){
	_ParagraphFormat para = m_sel.get_ParagraphFormat();
	para.put_Alignment(Alignment);
	m_sel.put_ParagraphFormat(para);
	para.ReleaseDispatch();
}
void SccWordApi::SetPageSetup(int LeftMargin /* = 20 */,int RightMargin /* = 20 */,int TopMargin /* = 20 */,int BottomMargin /* = 20 */){
	m_doc = m_app.get_ActiveDocument();
	PageSetup page = m_doc.get_PageSetup();

	// 设置页边距
	page.put_LeftMargin((float)LeftMargin);
	page.put_RightMargin((float)RightMargin);
	page.put_TopMargin((float)TopMargin);
	page.put_BottomMargin((float)BottomMargin);

	// 设置一页的行数
	page.put_LinesPage(44);
	
	page.ReleaseDispatch();
}
void SccWordApi::SetHeaderFooter(CString encoding,CString HeaderText /* =  */,CString FooterText /* = */ ){
	m_win = m_doc.get_ActiveWindow();
	m_pane = m_win.get_ActivePane();
	m_view = m_pane.get_View();

	SrcCodeScanner scanner;
	// 若为UTF-8编码
	if (encoding == "UTF-8")
	{
		USES_CONVERSION;
		// CStringA <-> LPCSTR <-> const char*
		HeaderText = scanner.GBKToUTF8(W2A(HeaderText)).c_str();
		FooterText = scanner.GBKToUTF8(W2A(FooterText)).c_str();
	} 

	// 程序本身默认为GBK编码

	// ********************设置页眉***********************
	m_view.put_SeekView(9); // wdSeekCurrentPageHeader = 9
	// 设置字体(宋体小五不加粗)
	m_font = m_sel.get_Font();
	m_font.put_Size((float)9);
	m_font.put_Name(_T("宋体"));
	//m_sel.put_Font(m_font); // 此处不需要put_Font
	// 设置对齐
	_ParagraphFormat para = m_sel.get_ParagraphFormat();
	para.put_Alignment(3); // 左对齐
	//m_sel.put_ParagraphFormat(para); // 此处不需要put_ParagraphFormat
	m_sel.TypeText(HeaderText);

	// ********************设置页脚***********************
	m_view.put_SeekView(10); // wdSeekCurrentPageFooter = 9
	para.put_Alignment(2); //右对齐
	m_sel.put_ParagraphFormat(para);
	m_sel.TypeText(FooterText);
	// 添加页码


	m_view.put_SeekView(0); // 关闭页眉页脚，返回主控文档
	para.ReleaseDispatch();
}



void SccWordApi::AppClose(){
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	if (!m_app.m_lpDispatch)
	{
		AfxMessageBox(_T("Word程序不存在，无法关闭!"));
	} 
	else
	{
		try // 用于解决"RPC服务器不可用"的BUG
		{
			m_app.Quit(vOpt,vOpt,vOpt);
		}
		catch (COleException* e)
		{
			throw e; // 将异常往上抛出，由OnBnClickedCancel() 捕获并处理
		}

		// 释放资源放在析构函数中

		//MessageBox(hWnd,_T("程序正在关闭..."),_T("代码扫描器"),MB_OK|MB_ICONINFORMATION);
	}
}