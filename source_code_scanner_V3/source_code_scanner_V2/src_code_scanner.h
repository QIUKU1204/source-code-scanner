#pragma once
#include <string>
#include <regex>
#include <vector>
#include <fstream>
#include "Word_Api.h"
using namespace std;

class SrcCodeScanner
{
public:
	SrcCodeScanner();
	~SrcCodeScanner();
public:
	/**
	*   去除字符串s 首尾的空字符、回车符、换行符及制表符
	*   @param[in] s              需要清洗的字符串
	**/
	void trim(string &s);

	/**
	*   将GBK编码的字符串转换为UTF-8编码
	*   @param[in] str_GBK        GBK编码的字符串
	**/
	string GBKToUTF8(const char* str_GBK);

	/**
	*   将UTF-8编码的字符串转换为GBK编码
	*   @param[in] str_UTF8       UTF-8编码的字符串
	**/
	string UTF8ToGBK(const char* str_UTF8);

	/**
	*   将宽字符串转换为普通的字符串
	*   @param[in] wchar          宽字符串
	*   @param[in] str            普通字符串
	**/
	void WCHARToString(WCHAR *wchar,string &str);

	/**
	*   替换字符串中的指定字符或子字符串
	*   @param[in] str            字符串
	*   @param[in] old_value      被替换的子字符串
	*   @param[in] new_value      进行替换的子字符串
	**/
	void ReplaceCharacters(string &str,string old_value = "\r\n",string new_value = "<br>");

	/**
	*   从文件夹中获取所有符合条件的文件
	*   @param[in] folder_path     文件夹路径
	*   @param[in] files_vc        存放文件路径的vector容器
	*   @param[in] file_count      计算放入容器的文件数量
	*   @param[in] file_type       文件类型
	**/
	void GetFilesFromFolder(string folder_path,vector<string>& files_vc,long& file_count,string file_type = "\\*");

	/**
	*   读取文件内容
	*   @param[in] filename         文件名路径
	*   @param[in] filedata         文件数据
	**/
	bool ReadFileData(string filename,string &filedata);

	/**
	*   对数据执行正则匹配
	*   @param[in] data             数据
	*   @param[in] pattern          正则表达式
    *   @param[in] vc               存放匹配结果的容器
	*   @param[in] position         匹配结果中的子字符串位置
	**/
	int RegexSearch(string data,regex pattern,vector<string> &vc,int position);

	/**
	*   从文件内容中提取类区域块
	*   @param[in] filename         文件名路径
	*   @param[in] class_block_vc   存放类区域块的容器
	**/
	bool GetClassBlock(string filename,vector<string> &class_block_vc);

	/**
	*   从类区域块数据中匹配提取注释信息并进行处理
	*   @param[in] class_block      类区域块
	*   @param[in] class_number     类区域块编号
	*   @param[in] complete_class   类说明 + 类名
	*   @param[in] xxxxxx_vc        存放注释信息的容器
	**/
	bool GetWantedData(string class_block,int class_number,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc,
				vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc);
	
	/**
	*   检查文件/文件夹路径容器
	*   @param[in] path_vc           文件/文件夹路径容器
	*   @param[in] hWnd              当前窗口句柄
	*   @param[in] file_extensions   需要扫描/过滤的文件类型/后缀
	**/
	bool CheckPathVector(vector<string> &path_vc,HWND hWnd,vector<string> file_extensions);

	/**
	*   根据提取到的注释信息生成Word 文档
	*   @param[in] filename          文件名路径
	*   @param[in] file_extensions   需要扫描/过滤的文件类型/后缀
	*   @param[in] encoding          生成文档的编码类型
	*   @param[in] wordOpt           Word 文档操作对象
	*   @param[in] header            页眉文本
	*   @param[in] footer            页脚文本
	**/
	void GenerateWordDoc(string filename,vector<string> file_extensions,string encoding,SccWordApi &wordOpt,CString header = "",CString footer = "");
	
	/**
	*   根据提取到的注释信息生成Markdown 文件
	*   @param[in] filename          文件名路径
	*   @param[in] file_extensions   需要扫描/过滤的文件类型/后缀
	*   @param[in] encoding          生成文档的编码类型
	*   @param[in] header            页眉文本
	*   @param[in] footer            页脚文本
	**/
	void GenerateMarkdownFile(string filename,vector<string> file_extensions,string encoding,string header = "GeoBeans平台二次开发库架构及主要接口说明文档",string footer = "北京中遥地网信息技术有限公司");

private:
	vector<string> class_block_vc;
	vector<string> class_name_vc;
	vector<string> class_desc_vc;
	vector<string> name_vc;
	vector<string> form_vc;
	vector<string> desc_vc;
	vector<string> param_vc;
	vector<string> return_vc;
	vector<string> example_vc;    // 存放匹配结果的容器
	string complete_class;

	string log_path;              // 日志文件路径
	ofstream logfile;
	char working_path[MAX_PATH];  // 当前工作路径
	char drive[_MAX_DRIVE];       // 磁盘名
	char dir[_MAX_DIR];           // 路径名
	char fname[_MAX_FNAME];       // 文件名
	char ext[_MAX_EXT];           // 后缀名

	SYSTEMTIME st;
	char str_time[128];           // 系统当前时间
}; 