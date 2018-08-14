#pragma once

class AFX_EXT_CLASS SccWordApi{
public:
	SccWordApi(void);
	~SccWordApi(void);

public:
	void ShowApp(BOOL flag);
	// ��COM���г�ʼ��
	BOOL InitCOM(); 	


	// ����һ��word����
	BOOL CreateApp();
	// ����һ��word�ĵ�
	BOOL CreateDocument();
	// ����һ��word�����word�ĵ�
	BOOL Create();
	// �رյ�ǰWord����
	void AppClose();


	// ��һ��word�ĵ�
	BOOL Open(CString FileName,BOOL ReadOnly = FALSE,BOOL AddToRecentFiles = FALSE);
	// �ر�һ��word�ĵ�
	BOOL Close(BOOL SaveChange = FALSE);
	// �����ĵ�
	BOOL Save();
	// ����Ϊָ������
	BOOL SaveAs(CString FileName,int SaveType = 0);


	///////////////////////�ļ�д����//////////////////////////////////////////////
	// д���ı�
	void WriteText(CString Text);
	// �س���N��
	void NewLine(int nCount = 1);
	// �س���N�к�д���ı�
	void WriteTextAfterNewLine(CString Text,int nCount = 1);
	// �ƶ���β��
	void EndLine();
	// ��ҳ
	void PageBreak();


	//////////////////////��������////////////////////////////////////////////////
	void SetFont(BOOL Bold,BOOL Italic = FALSE,BOOL UnderLine = FALSE);
	void SetFont(CString FontName,int FontSize = 9,long FontColor = 0,long Bold = 0);
	void SetTableFont(int Row,int Column,CString FontName,int FontSize = 9,long FontColor = 0,long FontBackColor = 0);


	//////////////////////������////////////////////////////////////////////////
	/**
	*   ��������ȷ���ı��
	*   @param[in] Row            ��
	*   @param[in] Column         ��
	**/
	void CreateTable(int Row,int Column);

	/**
	*   ���ı�����д��ָ����Ԫ����
	*   @param[in] Row            ��
	*   @param[in] Column         ��
	*   @param[in] Text           �ı�����
	**/
	void WriteCellText(int Row,int Column,CString Text);

	/**
	*   ���ñ�񱳾���ɫ
	*   @param[in] Row            ��
	*   @param[in] Column         ��
	*   @param[in] ShadingColor   ������ɫ
	**/
	void SetTableShading(int Row,int Column,long ShadingColor = 0);


	/////////////////////////��ǩ�������Ӳ���/////////////////////////////////////
	/**
	*	������ǩ
	*	@param[in] bmarkName ��ǩ����
	**/
	void CreateBookMark(CString bmarkName);

	/**
	*	�����ĵ��ڲ�������
	*	@param[in] hlinkName ����������
	*	@param[in] bmarkName ��ǩ����
	**/
	void CreateHyperLink(CString hlinkName,CString bmarkName);

	/**
	*	�ض�λ�ĵ��ڲ�������
	*	@param[in] hlinkName ����������
	*	@param[in] bmarkName ��ǩ����
	**/
	void RepositionHyperLink(CString hlinkName,CString bmarkName);    


	/////////////////////////////ͼƬ����/////////////////////////////////////////
	/**
	*	����ͼƬ
	*	@param[in] fileName ͼƬ·��
	**/
	void InsertShapes(CString fileName);

	/**
	*	����ͼƬ
	*	@param[in] pBitmap  bitmapͼƬ
	**/
	void InsertShapes(CBitmap *pBitmap);


	///////////////////////�����ĵ�����////////////////////////////////////////////
	/**
	*   ���ö�������
	*   @param[in] Alignment ���뷽ʽ
	*   - 1 ����
	*   - 2 �Ҷ���
	*   - 3 �����
	**/
	void SetParagraphFormat(int Alignment = 3);

	/**
	*   ����ҳ������
	*   @param[in] LeftMargin    ��߾�
	*   @param[in] RightMargin   �ұ߾�
	*   @param[in] TopMargin     �ϱ߾�
	*   @param[in] BottomMargin  �±߾�
	**/
	void SetPageSetup(int LeftMargin = 20,int RightMargin = 20,int TopMargin = 20,int BottomMargin = 20);

	/**
	*   ����ҳüҳ��
	*   @param[in] HeaderText   ҳü����
	*   @param[in] FooterText   ҳ������
	**/
	void SetHeaderFooter(CString encoding,CString HeaderText = "GeoBeansƽ̨���ο�����ܹ�����Ҫ�ӿ�˵���ĵ�",CString FooterText = "������ң������Ϣ�������޹�˾");

	/**
	*   �����Զ����
	**/
	void SetList();
};