#pragma once

class AFX_EXT_CLASS SccWordApi{
public:
	SccWordApi(void);
	~SccWordApi(void);

public:
	void ShowApp(BOOL flag);
	// 对COM进行初始化
	BOOL InitCOM(); 	


	// 创建一个word程序
	BOOL CreateApp();
	// 创建一个word文档
	BOOL CreateDocument();
	// 创建一个word程序和word文档
	BOOL Create();
	// 关闭当前Word程序
	void AppClose();


	// 打开一个word文档
	BOOL Open(CString FileName,BOOL ReadOnly = FALSE,BOOL AddToRecentFiles = FALSE);
	// 关闭一个word文档
	BOOL Close(BOOL SaveChange = FALSE);
	// 保存文档
	BOOL Save();
	// 保存为指定类型
	BOOL SaveAs(CString FileName,int SaveType = 0);


	///////////////////////文件写操作//////////////////////////////////////////////
	// 写入文本
	void WriteText(CString Text);
	// 回车换N行
	void NewLine(int nCount = 1);
	// 回车换N行后写入文本
	void WriteTextAfterNewLine(CString Text,int nCount = 1);
	// 移动到尾行
	void EndLine();
	// 分页
	void PageBreak();


	//////////////////////字体设置////////////////////////////////////////////////
	void SetFont(BOOL Bold,BOOL Italic = FALSE,BOOL UnderLine = FALSE);
	void SetFont(CString FontName,int FontSize = 9,long FontColor = 0,long Bold = 0);
	void SetTableFont(int Row,int Column,CString FontName,int FontSize = 9,long FontColor = 0,long FontBackColor = 0);


	//////////////////////表格操作////////////////////////////////////////////////
	/**
	*   创建行列确定的表格
	*   @param[in] Row            行
	*   @param[in] Column         列
	**/
	void CreateTable(int Row,int Column);

	/**
	*   将文本内容写入指定单元格内
	*   @param[in] Row            行
	*   @param[in] Column         列
	*   @param[in] Text           文本内容
	**/
	void WriteCellText(int Row,int Column,CString Text);

	/**
	*   设置表格背景样色
	*   @param[in] Row            行
	*   @param[in] Column         列
	*   @param[in] ShadingColor   背景颜色
	**/
	void SetTableShading(int Row,int Column,long ShadingColor = 0);


	/////////////////////////标签、超链接操作/////////////////////////////////////
	/**
	*	创建书签
	*	@param[in] bmarkName 书签名称
	**/
	void CreateBookMark(CString bmarkName);

	/**
	*	创建文档内部超链接
	*	@param[in] hlinkName 超链接名称
	*	@param[in] bmarkName 书签名称
	**/
	void CreateHyperLink(CString hlinkName,CString bmarkName);

	/**
	*	重定位文档内部超链接
	*	@param[in] hlinkName 超链接名称
	*	@param[in] bmarkName 书签名称
	**/
	void RepositionHyperLink(CString hlinkName,CString bmarkName);    


	/////////////////////////////图片操作/////////////////////////////////////////
	/**
	*	插入图片
	*	@param[in] fileName 图片路径
	**/
	void InsertShapes(CString fileName);

	/**
	*	插入图片
	*	@param[in] pBitmap  bitmap图片
	**/
	void InsertShapes(CBitmap *pBitmap);


	///////////////////////设置文档属性////////////////////////////////////////////
	/**
	*   设置对齐属性
	*   @param[in] Alignment 对齐方式
	*   - 1 居中
	*   - 2 右对齐
	*   - 3 左对齐
	**/
	void SetParagraphFormat(int Alignment = 3);

	/**
	*   设置页面属性
	*   @param[in] LeftMargin    左边距
	*   @param[in] RightMargin   右边距
	*   @param[in] TopMargin     上边距
	*   @param[in] BottomMargin  下边距
	**/
	void SetPageSetup(int LeftMargin = 20,int RightMargin = 20,int TopMargin = 20,int BottomMargin = 20);

	/**
	*   设置页眉页脚
	*   @param[in] HeaderText   页眉内容
	*   @param[in] FooterText   页脚内容
	**/
	void SetHeaderFooter(CString encoding,CString HeaderText = "GeoBeans平台二次开发库架构及主要接口说明文档",CString FooterText = "北京中遥地网信息技术有限公司");

	/**
	*   设置自动编号
	**/
	void SetList();
};