#include "stdafx.h"
#include "Word_Api.h"
#include "src_code_scanner.h"


_Application    m_app;           ///< ����word
Documents       m_docs;          ///< word�ĵ�����
_Document       m_doc;           ///< һ��word�ļ�
_Font           m_font;          ///< �������
Selection       m_sel;           ///< ѡ��༭����û�ж����ʱ����ǲ����
Tables          m_tabs;
Table           m_tab;           ///< ������
Bookmarks       m_bmarks;        
Bookmark        m_bmark;         ///< ��ǩ����
Hyperlinks      m_hlinks;
Hyperlink       m_hlink;         ///< �����Ӷ���
InlineShapes    m_ishapes;       
InlineShape     m_ishape;        ///< ͼƬ����
Borders         m_bds;
Border          m_bd;            ///< ���߿����
Window			m_win;           ///< ���ڶ���
Pane			m_pane;			 ///< Pane(����)����
View			m_view;			 ///< ��ͼ����
HeaderFooter	m_hefo;			 ///< ҳüҳ�Ŷ���
List			m_list;          ///< List����
Field           m_field;         
Fields			m_fields;		 ///< �����


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
	// ��ʼ��COM
	if (!AfxOleInit())
	{
		AfxMessageBox(_T("�޷���ʼ��COM�Ķ�̬���ӿ�"));
		return FALSE;
	}
	else{
		AfxMessageBox(_T("COM��ʼ���ɹ���"));
		return TRUE;
	}
}

BOOL SccWordApi::CreateApp(){
	if (!m_app.CreateDispatch(_T("Word.Application")))
	{
		AfxMessageBox(_T("���񴴽�ʧ�ܣ���ȷ���Ѿ���װ��Office 2000�����ϰ汾"));
		return FALSE;
	} 
	else
	{
		//m_app.put_Visible(TRUE);

		//AfxMessageBox(_T("Word�����ɹ�����ȷ�ϼ�����"),MB_OK|MB_ICONINFORMATION);

		return TRUE;
	}
}

BOOL SccWordApi::CreateDocument(){
	if (!m_app.m_lpDispatch)
	{
		AfxMessageBox(_T("û�д�Word�����ĵ�����ʧ�ܣ�"), MB_OK | MB_ICONWARNING);
		return FALSE;
	} 
	else
	{
		try // ���ڽ��"RPC������������"��BUG
		{
			m_docs = m_app.get_Documents();
		}
		catch (COleException* e)
		{
			throw e; // ���쳣�����׳�����GenerateWordDoc() ���񲢴���
		}
		
		if (m_docs.m_lpDispatch == NULL)
		{
			AfxMessageBox(_T("Word�ĵ�������ʧ�ܣ�"));
			return FALSE;
		} 
		else
		{
			COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
			m_doc = m_docs.Add(vOpt,vOpt,vOpt,vOpt);

			if (!m_doc.m_lpDispatch)
			{
				AfxMessageBox(_T("Word�ĵ�����ʧ�ܣ�"));
				return FALSE;
			} 
			else
			{
				m_sel = m_app.get_Selection(); // ��õ�ǰWord����
				if (!m_sel.m_lpDispatch)
				{
					AfxMessageBox(_T("ѡ��༭����Selection��ȡʧ�ܣ�"));
					return FALSE;
				} 
				else
				{
					//AfxMessageBox(_T("Word�ĵ������ɹ���"));
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
			AfxMessageBox(_T("Word�ĵ�������ʧ�ܣ�"));
			return FALSE;
		}
	}
	m_doc = m_docs.Open(&Name,varFalse,&Read,&AddToR,vOpt,vOpt,
		vFalse,vOpt,vOpt,&format,vOpt,vTrue,vOpt,vOpt,vOpt,vOpt);
	if (!m_doc.m_lpDispatch)
	{
		AfxMessageBox(_T("Word�ĵ�����ʧ�ܣ�"));
		return FALSE;
	} 
	else
	{
		m_sel = m_app.get_Selection();
		if (!m_sel.m_lpDispatch)
		{
			AfxMessageBox(_T("Word�ĵ���ʧ�ܣ�"));
			return FALSE;
		}
		//AfxMessageBox(_T("Word�ĵ��򿪳ɹ���"));
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
		AfxMessageBox(_T("Word�ĵ��ر�ʧ�ܣ�"));
		return FALSE;
	}
	else
	{
		if(TRUE==SaveChange)
		{
			Save();
		}
		//�����һ��������vTrue ����ִ��󣬿����Ǻ���Ĳ���ҲҪ��Ӧ�ı仯
		//��vba û�и���Ӧ���� �Ҿ������ַ���������
		m_doc.Close(&vFalse,&vOpt,&vOpt);
		//AfxMessageBox(_T("Word�ĵ��رճɹ���"));
	}
	return TRUE;
}

BOOL SccWordApi::Save(){
	if (!m_doc.m_lpDispatch)
	{
		AfxMessageBox(_T("û�д���Document���󣬱���ʧ��"));
		return FALSE;
	} 
	else
	{
		m_doc.Save();
		//AfxMessageBox(_T("����ɹ���"));
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
		AfxMessageBox(_T("û�д���Document���󣬱���ʧ��"));
		return FALSE;
	} 
	else
	{
		// ��ð��պ�����д����Ȼ���ܳ���
		m_doc.SaveAs(&cFileName,&FileFormat,&vFalse,COleVariant(_T("")),&vTrue,
			COleVariant(_T("")),&vFalse,&vFalse,&vFalse,&vFalse,&vFalse,&vOpt,&vOpt,&vOpt,&vOpt,&vOpt);
		//AfxMessageBox(_T("Word�ĵ�����ɹ���"));
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
			m_sel.TypeParagraph(); // �½�һ��
		}
	}
}
void SccWordApi::WriteTextAfterNewLine(CString Text,int nCount /* = 1 */){
	NewLine(nCount);
	WriteText(Text);
}
void SccWordApi::EndLine(){
	m_sel = m_app.get_Selection();
	m_sel.EndKey(COleVariant((short)6),COleVariant((short)0)); // ������ƶ������һ��
}
void SccWordApi::PageBreak(){
	m_sel = m_app.get_Selection();
	m_sel.InsertBreak(COleVariant((long)7)); // �����ҳ��
}



void SccWordApi::SetFont(BOOL Bold,BOOL Italic /* = FALSE */,BOOL UnderLine /* = FALSE */){
	if (!m_sel.m_lpDispatch)
	{
		AfxMessageBox(_T("ѡ��༭����Ϊ�գ��޷���������"));
		return;
	} 
	else
	{
		// const char *���͵�ʵ����LPCTSTR���͵��вβ�����
		// m_sel.put_Text("F");
		m_sel.put_Text(_T("F")); 
		m_font = m_sel.get_Font(); // ����������
		m_font.put_Bold(Bold);  
		m_font.put_Italic(Italic);
		m_font.put_Underline(UnderLine); // ��������ĸ������
		m_sel.put_Font(m_font); //�����������ò���Ч
	}
}
void SccWordApi::SetFont(CString FontName,int FontSize /* = 9 */,long FontColor /* = 0 */,long Bold /* = 0 */){
	if (!m_sel.m_lpDispatch)
	{
		AfxMessageBox(_T("ѡ��༭����Ϊ�գ��޷���������"));
		return;
	}
	// ʹ�� get_Font ��ȡ�������
	// ʹ�� put_Text ��ȡ��������
	m_sel.put_Text(_T("a"));
	m_font = m_sel.get_Font(); // ��ȡ�������
	m_font.put_Size((float)FontSize);
	m_font.put_Name(FontName);
	m_font.put_Color(FontColor);
	m_font.put_Bold(Bold);
	m_sel.put_Font(m_font); // ѡ���������
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
		AfxMessageBox(_T("����������ʧ��"));
		return;
	} 
	else
	{
		m_tabs.Add(m_sel.get_Range(),Row,Column,&DefaultBehavior,&AutoFitBehavior);
		m_tab = m_tabs.Item(m_tabs.get_Count()); // 

		// ���ñ��Ŀ��(�ٷֱ�%)
		m_tab.put_PreferredWidth(101.5);

		// ���ñ��ı߿�����(��Ҫ֪����ֵ)
		/*m_bds = m_tab.get_Borders();
		m_bds.put_OutsideLineStyle(20);
		m_bds.put_OutsideColorIndex(0);
		m_bds.put_InsideLineStyle(10);
		m_bds.put_InsideColorIndex(0);*/
	}
}
void SccWordApi::WriteCellText(int Row,int Column,CString Text){
	Cell cell = m_tab.Cell(Row,Column); // �ӱ���л�ȡָ����Ԫ��
	cell.Select(); // ѡ��õ�Ԫ����Ϊ�༭����
	m_sel.TypeText(Text); // ���ı�����д��õ�Ԫ��
}
void SccWordApi::SetTableShading(int Row,int Column,long ShadingColor /* = 0 */){
	Cell cell = m_tab.Cell(Row,Column);
	cell.Select();
	Shading	shading = cell.get_Shading(); // ��ָ����Ԫ���ȡ��Shading(������ɫ)����
	shading.put_BackgroundPatternColor(ShadingColor);
}



void SccWordApi::CreateBookMark(CString bmarkName)
{
	m_doc = m_app.get_ActiveDocument();
	m_sel = m_app.get_Selection();
	m_bmarks=m_doc.get_Bookmarks(); 

	if(!m_bmarks.m_lpDispatch)
	{
		AfxMessageBox(_T("������ǩ����ʧ��"));
		return;
	}
	else
	{
		COleVariant vOpt;
		Range aRange = m_sel.get_Range();
		aRange.m_lpDispatch->AddRef();				///< ������仰���
		vOpt.vt = VT_DISPATCH;						//ָ��vt������
		vOpt.pdispVal = aRange;						//��range�����m_lpDispatch��ֵ��vt
		m_bmark = m_bmarks.Add(bmarkName, &vOpt);	//����һ����bmarkName����ǩ
	}
}
void SccWordApi::CreateHyperLink(CString hlinkName,CString bmarkName)
{
	COleVariant     vAddress(_T("")),vSubAddress(bmarkName),vScreenTip(_T("")),vTextToDisplay(hlinkName);

	m_doc=m_app.get_ActiveDocument();
	Range aRange = m_sel.get_Range();

	m_hlinks = m_doc.get_Hyperlinks();
	m_hlinks.Add(
		aRange,				//Object�����衣ת��Ϊ�����ӵ��ı���ͼ�Ρ�
		vAddress,			//Variant ���ͣ���ѡ��ָ�������ӵĵ�ַ���˵�ַ�����ǵ����ʼ���ַ��Internet ��ַ���ļ�������ע�⣬Microsoft Word �����õ�ַ����ȷ�ԡ�
		vSubAddress,		//Variant ���ͣ���ѡ��Ŀ���ļ��ڵ�λ����������ǩ���������������õ�Ƭ��š�
		vScreenTip,			//Variant ���ͣ���ѡ�������ָ�����ָ���ĳ�������ʱ��ʾ�Ŀ���������Ļ��ʾ�����ı���Ĭ��ֵΪ Address��
		vTextToDisplay,		//Variant ���ͣ���ѡ��ָ���ĳ����ӵ���ʾ�ı����˲�����ֵ��ȡ���� Anchor ָ�����ı���ͼ�Ρ�
		vSubAddress			//Variant ���ͣ���ѡ��Ҫ�����д�ָ���ĳ����ӵĿ�ܻ򴰿ڵ����֡�

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

	// ����ҳ�߾�
	page.put_LeftMargin((float)LeftMargin);
	page.put_RightMargin((float)RightMargin);
	page.put_TopMargin((float)TopMargin);
	page.put_BottomMargin((float)BottomMargin);

	// ����һҳ������
	page.put_LinesPage(44);
	
	page.ReleaseDispatch();
}
void SccWordApi::SetHeaderFooter(CString encoding,CString HeaderText /* =  */,CString FooterText /* = */ ){
	m_win = m_doc.get_ActiveWindow();
	m_pane = m_win.get_ActivePane();
	m_view = m_pane.get_View();

	SrcCodeScanner scanner;
	// ��ΪUTF-8����
	if (encoding == "UTF-8")
	{
		USES_CONVERSION;
		// CStringA <-> LPCSTR <-> const char*
		HeaderText = scanner.GBKToUTF8(W2A(HeaderText)).c_str();
		FooterText = scanner.GBKToUTF8(W2A(FooterText)).c_str();
	} 

	// ������Ĭ��ΪGBK����

	// ********************����ҳü***********************
	m_view.put_SeekView(9); // wdSeekCurrentPageHeader = 9
	// ��������(����С�岻�Ӵ�)
	m_font = m_sel.get_Font();
	m_font.put_Size((float)9);
	m_font.put_Name(_T("����"));
	//m_sel.put_Font(m_font); // �˴�����Ҫput_Font
	// ���ö���
	_ParagraphFormat para = m_sel.get_ParagraphFormat();
	para.put_Alignment(3); // �����
	//m_sel.put_ParagraphFormat(para); // �˴�����Ҫput_ParagraphFormat
	m_sel.TypeText(HeaderText);

	// ********************����ҳ��***********************
	m_view.put_SeekView(10); // wdSeekCurrentPageFooter = 9
	para.put_Alignment(2); //�Ҷ���
	m_sel.put_ParagraphFormat(para);
	m_sel.TypeText(FooterText);
	// ���ҳ��


	m_view.put_SeekView(0); // �ر�ҳüҳ�ţ����������ĵ�
	para.ReleaseDispatch();
}



void SccWordApi::AppClose(){
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	if (!m_app.m_lpDispatch)
	{
		AfxMessageBox(_T("Word���򲻴��ڣ��޷��ر�!"));
	} 
	else
	{
		try // ���ڽ��"RPC������������"��BUG
		{
			m_app.Quit(vOpt,vOpt,vOpt);
		}
		catch (COleException* e)
		{
			throw e; // ���쳣�����׳�����OnBnClickedCancel() ���񲢴���
		}

		// �ͷ���Դ��������������

		//MessageBox(hWnd,_T("�������ڹر�..."),_T("����ɨ����"),MB_OK|MB_ICONINFORMATION);
	}
}