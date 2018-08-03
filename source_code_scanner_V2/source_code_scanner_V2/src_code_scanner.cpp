#include "stdafx.h"
#include "src_code_scanner.h"
#include "Word_Api.h"
#include <regex>
#include <iostream>
#include <fstream>
#include <streambuf>
#include <string>
#include <sstream>
#include <queue>
#include <vector>
#include <Windows.h>
using namespace std;


SrcCodeScanner::SrcCodeScanner(){}

SrcCodeScanner::~SrcCodeScanner(){}

void SrcCodeScanner::trim(string &s)
{
	size_t n = s.find_last_not_of( " \r\n\t" );
	if( n != string::npos )
	{
		s.erase( n + 1 , s.size() - n );
	}

	n = s.find_first_not_of ( " \r\n\t" );
	if( n != string::npos )
	{
		s.erase( 0 , n );
	}
}

string SrcCodeScanner::GBKToUTF8(const char* str_GBK){
	int len = MultiByteToWideChar(CP_ACP, 0, str_GBK, -1, NULL, 0);
	wchar_t* w_str = new wchar_t[len+1];
	memset(w_str, 0, len+1);
	MultiByteToWideChar(CP_ACP, 0, str_GBK, -1, w_str, len);
	len = WideCharToMultiByte(CP_UTF8, 0, w_str, -1, NULL, 0, NULL, NULL);
	char* str = new char[len+1];
	memset(str, 0, len+1);
	WideCharToMultiByte(CP_UTF8, 0, w_str, -1, str, len, NULL, NULL);
	string str_Temp = str;
	if(w_str) delete[] w_str;
	if(str) delete[] str;
	return str_Temp;
}

string SrcCodeScanner::UTF8ToGBK(const char* str_UTF8){
	int len = MultiByteToWideChar(CP_UTF8, 0, str_UTF8, -1, NULL, 0);
	wchar_t* w_sz_GBK = new wchar_t[len+1];
	memset(w_sz_GBK, 0, len*2+2);
	MultiByteToWideChar(CP_UTF8, 0, str_UTF8, -1, w_sz_GBK, len);
	len = WideCharToMultiByte(CP_ACP, 0, w_sz_GBK, -1, NULL, 0, NULL, NULL);
	char* sz_GBK = new char[len+1];
	memset(sz_GBK, 0, len+1);
	WideCharToMultiByte(CP_ACP, 0, w_sz_GBK, -1, sz_GBK, len, NULL, NULL);
	string str_Temp(sz_GBK);
	if(w_sz_GBK) delete[] w_sz_GBK;
	if(sz_GBK) delete[] sz_GBK;
	return str_Temp;

}

void SrcCodeScanner::WCHARToString(WCHAR *wchar,string &str){
	wchar_t * wideText = wchar;
	// 第一次调用返回转换后的字符串长度
	int len = WideCharToMultiByte(CP_OEMCP,NULL,wideText,-1,NULL,0,NULL,FALSE);
	char *multiText; // char*类型临时数组，作为赋值的中间变量
	multiText = new char[len];
	// 第二次调用将 宽字符字符串 转换为 多字节字符串
	WideCharToMultiByte (CP_OEMCP,NULL,wideText,-1,multiText,len,NULL,FALSE);
	str = multiText;// char*直接赋值给string
	delete []multiText;
}

bool SrcCodeScanner::ReadFileData(string filename,string &data){
	regex ftype_pattern("[^/*""<>|?]+?\\.h");
	if (!regex_match(filename,ftype_pattern))
	{
		AfxMessageBox(_T("文件类型错误，请选择C++头文件(.h)！"));
		return false;
	}
	ifstream in(filename,ios::binary);
	string data_tmp((istreambuf_iterator<char>(in)),
					istreambuf_iterator<char>());
	data = data_tmp;
	// data为空的处理
	if (data.empty()) // or data.length == 0
	{
		AfxMessageBox(_T("找不到指定的文件！"));
		return false;
	}
	return true;
}

bool SrcCodeScanner::RegexSearch(string data,regex pattern,vector<string> &vc,int position){
	// 使用类 regex_iterator 进行多次匹配，返回所有符合的匹配结果
	auto words_begin = sregex_iterator(data.begin(),data.end(),pattern);
	auto words_end = sregex_iterator();

	// 定长的字符串数组，需要动态分配内存空间
	//string * matches;
	int size = 0;
	for (sregex_iterator it = words_begin; it != words_end; ++it)
	{
		vc.push_back(it->format("$" + to_string((long long)position)));
		++size;
	}
	// 当 size = 0 时的处理
	if (size == 0)
	{
		return false;
	}
	return true;

	//matches = new string[size];
	//delete [] matches;
}

void SrcCodeScanner::GenerateWordDoc(string filename,SccWordApi &wordOpt){
	/////////////////////////文件读写与正则表达式//////////////////////////

	// 读取文件内容(UTF8)
	string filedata;
	if (!ReadFileData(filename,filedata))
	{
		return;
	}

	// 构建正则表达式(需要优化)(ToUTF8)
	regex class_name_pattern(GBKToUTF8("Class ([A-Za-z: ]+) : Begin."));
	regex class_desc_pattern(GBKToUTF8("类说明：(.+)")); // :与：
	regex name_pattern(GBKToUTF8("名称:([A-Za-z~ ]+)"));
	regex form_pattern(GBKToUTF8("(接口形式:.+)"));
	regex desc_pattern(GBKToUTF8("(功能描述:.+)"));
	regex param_pattern(GBKToUTF8("// 参数说明:((.|\\r|\\n)+?)// 返回值"));
	regex return_pattern(GBKToUTF8("// 返回值:((.|\\r|\\n)+?)// 示例"));
	regex example_pattern(GBKToUTF8("// 示例:((.|\\r|\\n)+?)//"));

	// 存放正则匹配结果的容器
	vector<string> class_name_vc;
	vector<string> class_desc_vc;
	vector<string> name_vc;
	vector<string> form_vc;
	vector<string> desc_vc;
	vector<string> param_vc;
	vector<string> return_vc;
	vector<string> example_vc;

	// 执行匹配(UTF8)
	bool flag = RegexSearch(filedata,class_name_pattern,class_name_vc,1);
	flag = RegexSearch(filedata,class_desc_pattern,class_desc_vc,1);
	flag = RegexSearch(filedata,name_pattern,name_vc,1);
	flag = RegexSearch(filedata,form_pattern,form_vc,1);
	flag = RegexSearch(filedata,desc_pattern,desc_vc,1);
	flag = RegexSearch(filedata,param_pattern,param_vc,1);
	flag = RegexSearch(filedata,return_pattern,return_vc,1);
	flag = RegexSearch(filedata,example_pattern,example_vc,1);
	if (!flag)
	{
		// 弹窗提示具体是哪个文件匹配失败，没有生成Word文档
		string warning = filename + "匹配失败，请查看其注释格式！";
		AfxMessageBox((CString)warning.c_str());
		return;
	}

	// 对匹配结果进行处理(数字序号待完善)

	string complete_class = GBKToUTF8("1.1.1 ") + class_desc_vc[0] + class_name_vc[0];

	for (unsigned int i = 0;i < name_vc.size();i++)
	{
		// 处理接口名称(ToUTF8)
		string str_number = std::to_string((long long)(i + 1));
		name_vc[i] = GBKToUTF8("1.1.1 ") + str_number + name_vc[i];

		// 处理接口描述(ToUTF8)

		// 处理接口参数(ToUTF8)
		string str_tmp = param_vc[i]; // 两个不同的迭代器同时操作同一个vector会出错，因此赋值给临时string类型变量
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);
		param_vc[i] = GBKToUTF8("参数: ") + str_tmp;

		// 处理接口返回值(ToUTF8)
		str_tmp = return_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end()); // 将导致 -1 变为 1
		trim(str_tmp);
		return_vc[i] = GBKToUTF8("返回值: ") + str_tmp;

		// 处理接口示例代码(ToUTF8)
		str_tmp = example_vc[i];
		trim(str_tmp);
		example_vc[i] = GBKToUTF8("示例代码: ") + str_tmp;
	}


	////////////////操作Word(生成的Word文档默认为GBK编码)//////////////////
	//CsccWordApi wordOpt;
	//wordOpt.CreateApp();
	wordOpt.CreateDocument();

	wordOpt.SetPageSetup(54,54,72,72);
	wordOpt.SetParagraphFormat(3);
	wordOpt.SetHeaderFooter();

	// 写入类名及类说明(ToGBK)
	wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  四号加粗
	wordOpt.WriteText(UTF8ToGBK(complete_class.c_str()).c_str()); 
	wordOpt.NewLine();

	// 循环写入每个接口函数的注释信息(ToGBK)
	for (unsigned int i = 0;i < name_vc.size();i++)
	{
		// 接口函数名
		wordOpt.SetFont(_T("Times New Roman"),12,0,1); // 小四加粗
		wordOpt.WriteText(UTF8ToGBK(name_vc[i].c_str()).c_str());
		wordOpt.NewLine();

		wordOpt.SetFont(_T("Times New Roman"),12,0,0); // 小四不加粗
		wordOpt.CreateTable(5,1);

		// 接口详细信息
		wordOpt.WriteCellText(1,1,UTF8ToGBK(form_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(2,1,UTF8ToGBK(desc_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(3,1,UTF8ToGBK(param_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(4,1,UTF8ToGBK(return_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(5,1,UTF8ToGBK(example_vc[i].c_str()).c_str());

		// 设置单元格背景颜色
		wordOpt.SetTableShading(1,1,16777215);
		wordOpt.SetTableShading(2,1,12500670);
		wordOpt.SetTableShading(3,1,16777215);
		wordOpt.SetTableShading(4,1,12500670);
		wordOpt.SetTableShading(5,1,16777215);

		wordOpt.EndLine();
	}

	// 保存Word文档、关闭Word程序
	regex filename_pattern("(.+?)\\.h");
	vector<string> filename_vc;
	RegexSearch(filename,filename_pattern,filename_vc,1);
	string saved_doc_name = filename_vc[0] + ".doc";
	// NOTE: string.c_str()可以赋值给CString，CString可以赋值给LPCTSTR，
	//       但是string.c_str()不可以直接赋值给LPCTSTR；
	//AfxMessageBox((CString)saved_doc_name.c_str()); // CStringW -> LPCTSTR ; 
	wordOpt.SaveAs(saved_doc_name.c_str()); // string.c_str() -> CString;
	wordOpt.Close(); // 关闭当前文档，但不关闭Word程序
}