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
#include <io.h>
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

void SrcCodeScanner::ReplaceCharacters(string &str,string old_value /* =  */,string new_value /* = */ ){
	int sign = str.find(old_value,0); // 从下标为0的位置开始匹配old_value，若匹配则返回首字母的下标
	while (sign != string::npos)
	{   
		// 从下标为sign的位置开始，使用new_value替换size个字符
		str.replace(sign,new_value.size(),new_value); // 此处最好是对等替换，不对等替换可能导致数据丢失 
		sign = str.find(old_value,sign + 1);
	}
	return;
}

void SrcCodeScanner::GetFilesFromFolder(string folder_path,vector<string>& files_vc,long& file_count,string file_type /* = */ )  
{  
	// 文件句柄  
	long hFile = 0;  
	// 文件信息  
	struct _finddata_t fileinfo;  
	string path_begin; 
	// 开始查找
	if((hFile = _findfirst(path_begin.assign(folder_path).append("\\*").c_str(),&fileinfo)) != -1)  
	{  
		do{  
			// 如果是目录,迭代之   
			if((fileinfo.attrib & _A_SUBDIR))  
			{  
				if(strcmp(fileinfo.name,".") != 0 && strcmp(fileinfo.name,"..") != 0)
					// 递归调用 GetFilesFromFolder
					GetFilesFromFolder(path_begin.assign(folder_path).append("\\").append(fileinfo.name),files_vc,file_count,file_type);  
			}  
			else  // 如果是file_type类型的文件,则加入容器 
			{  
				if (strstr(fileinfo.name,file_type.c_str()) != NULL)
				{
					files_vc.push_back(path_begin.assign(folder_path).append("\\").append(fileinfo.name));
					++file_count;
				}
			}  
		}while(_findnext(hFile, &fileinfo)  == 0); // 继续查找
		// 结束查找
		_findclose(hFile);  
	}  
}

bool SrcCodeScanner::ReadFileData(string filename,string &data){
	regex ftype_pattern("[^/*""<>|?]+?\\.h");
	if (!regex_match(filename,ftype_pattern)) // 判断文件是否是C++头文件(.h)
	{
		string warning = filename + "文件类型错误，请务必选择C++头文件(.h)！";
		//MessageBox(NULL,(CString)warning.c_str(),_T("提示信息"),MB_OK|MB_ICONINFORMATION); // 需传入窗口句柄
		AfxMessageBox((CString)warning.c_str(),MB_OK|MB_ICONINFORMATION);
		return false;
	}

	ifstream in(filename,ios::binary|ios::in);
	string data_tmp((istreambuf_iterator<char>(in)),istreambuf_iterator<char>());
	data = UTF8ToGBK(data_tmp.c_str()); // 重要的一步: UTF8 -> GBK ;

	// data为空的处理
	if (data.empty()) // or data.length == 0
	{
		//MessageBox(NULL,_T("找不到指定的文件！"),_T("提示信息"),MB_OK|MB_ICONINFORMATION); // 需传入窗口句柄
		AfxMessageBox(_T("找不到指定的文件！"),MB_OK|MB_ICONINFORMATION);
		return false;
	}
	return true;
}

int SrcCodeScanner::RegexSearch(string data,regex pattern,vector<string> &vc,int position){
	// 使用类 regex_iterator 进行多次匹配，返回所有符合的匹配结果
	auto words_begin = sregex_iterator(data.begin(),data.end(),pattern);
	auto words_end = sregex_iterator();

	for (sregex_iterator it = words_begin; it != words_end; ++it)
	{
		vc.push_back(it->format("$" + to_string((long long)position)));
	}

	if (vc.size() == 0) // 匹配失败返回1
	{
		return 1;
	}

	return 0; // 匹配成功返回0
}

bool SrcCodeScanner::GetWantedData(string filename,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc, 
						vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc){
	// 读取文件内容
	string filedata;
	if (!ReadFileData(filename,filedata)) 
	{
		return false; // 文件读取失败
	}

	// 构建正则表达式:兼顾匹配精度与匹配容错率
    // 优化方案:提供接口供用户输入自定义的正则表达式
	regex class_name_pattern("Class[ ]*([A-Za-z]+)[ ]*:[ ]*Begin\\.*"); 
	regex class_desc_pattern("类说明:*：*(.+)");      // 此处不能使用[:：]{1}
	regex name_pattern("名称[:：]{1}[ ]*([A-Za-z~ ]+)"); // 此处不能使用:*：*
	regex form_pattern("(接口形式[:：]{1}[ ]*.+)");
	regex desc_pattern("(功能描述[:：]{1}[ ]*.+)");
	regex param_pattern("//[ ]*参数说明[:：][ ]*((.|\\r|\\n)+?)//[ ]*返回值");
	regex return_pattern("//[ ]*返回值[:：][ ]*((.|\\r|\\n)+?)//[ ]*示例"); // 需要优化
	regex example_pattern("//[ ]*示例[:：][ ]*((.|\\r|\\n)+?)//");          // 需要优化

	// 执行匹配（成功返回0，失败返回1）
	int flag = RegexSearch(filedata,class_name_pattern,class_name_vc,1);
	flag += RegexSearch(filedata,class_desc_pattern,class_desc_vc,1);
	flag += RegexSearch(filedata,name_pattern,name_vc,1);
	flag += RegexSearch(filedata,form_pattern,form_vc,1);
	flag += RegexSearch(filedata,desc_pattern,desc_vc,1);
	flag += RegexSearch(filedata,param_pattern,param_vc,1);
	flag += RegexSearch(filedata,return_pattern,return_vc,1);
	flag += RegexSearch(filedata,example_pattern,example_vc,1);
	if (flag != 0)
	{
		// 弹窗提示具体是哪个文件匹配失败
		string warning = filename + "匹配失败，请查看其注释格式！";
		//MessageBox(NULL,(CString)warning.c_str(),_T("提示信息"),MB_OK|MB_ICONINFORMATION); // 需传入窗口句柄
		AfxMessageBox((CString)warning.c_str(),MB_OK|MB_ICONINFORMATION);
		return false;
	}

	// 对匹配结果进行处理
	complete_class = "1.1.1 " + class_desc_vc[0] + class_name_vc[0];
	for (size_t i = 0;i < name_vc.size();i++)
	{
		// 处理接口名称
		string str_number = std::to_string((long long)(i + 1));
		name_vc[i] = "1.1.1." + str_number + " " + name_vc[i];

		// 处理接口描述

		// 处理接口描述

		// 处理接口参数
		string str_tmp = param_vc[i]; // 两个不同的迭代器同时操作同一个vector会出错，因此赋值给临时string类型变量
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);
		param_vc[i] = "参数: " + str_tmp;

		// 处理接口返回值
		str_tmp = return_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end()); // 将导致 -1 变为 1
		trim(str_tmp);
		return_vc[i] = "返回值: " + str_tmp;

		// 处理接口示例代码
		str_tmp = example_vc[i];
		trim(str_tmp);
		example_vc[i] = "示例代码: " + str_tmp;
	}
	return true;
}

bool SrcCodeScanner::CheckPathVector(vector<string> &path_vc,HWND hWnd){ // 检查存放文件/文件夹路径的vector容器
	if (path_vc.size() == 0) // 检查容器是否为空，以确认是否选择了文件或文件夹
	{
		MessageBox(hWnd,_T("请先选择文件/文件夹！"),_T("提示信息"),MB_OK | MB_ICONINFORMATION);
		return false;
	}
	// 检查vc容器中的元素，若为文件夹，则将该文件夹下的所有头文件放入容器中
	for (size_t i = 0;i < path_vc.size();i++)
	{
		if (path_vc[i].find('.') == string::npos)
		{
			long file_count = 0;
			string folder_path = path_vc[i];
			path_vc.erase(path_vc.begin() + i);   // 删除文件夹路径
			string file_type = ".h"; // 优化方案: 在程序界面设置文件类型可选
			GetFilesFromFolder(folder_path,path_vc,file_count,file_type);
			if (file_count == 0){ 
				string warning = "文件夹" + folder_path + "为空！";
				MessageBox(hWnd,(CString)warning.c_str(),_T("提示信息"),MB_OK|MB_ICONINFORMATION);
			}
			i--; // 删除vector中的元素后，容器的size将减一，剩下的元素将整体前移一位
		}
	}
	if (path_vc.size() == 0) // 再次检查容器是否为空
	{
		return false;
	}
	return true;
}

void SrcCodeScanner::GenerateWordDoc(string filename,SccWordApi &wordOpt){
	//////////////////获取生成Word文档所需的处理好的数据//////////////////////
	// 定义存放正则匹配结果的容器
	vector<string> class_name_vc;
	vector<string> class_desc_vc;
	vector<string> name_vc;
	vector<string> form_vc;
	vector<string> desc_vc;
	vector<string> param_vc;
	vector<string> return_vc;
	vector<string> example_vc;
	string complete_class;

	// 读取(头)文件内容 -> 执行正则匹配 -> 处理匹配结果 -> 获得需要的处理好的数据
	if (!GetWantedData(filename,complete_class,class_name_vc,class_desc_vc,name_vc,form_vc,desc_vc,param_vc,return_vc,example_vc))
	{
		return; // 若没有获取写入所需的数据，则不再执行生成Word文档的操作
	}
	

	////////////////操作Word接口生成GBK编码的Word文档//////////////////
	wordOpt.CreateDocument();
	wordOpt.SetPageSetup(54,54,72,72);
	wordOpt.SetParagraphFormat(3);
	wordOpt.SetHeaderFooter();

	// 写入类名及类说明
	wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  四号加粗
	wordOpt.WriteText(complete_class.c_str()); 
	wordOpt.NewLine();

	// 循环写入每个接口函数的注释信息
	for (size_t i = 0;i < name_vc.size();i++)
	{
		// 接口函数名
		wordOpt.SetFont(_T("Times New Roman"),12,0,1); // 小四加粗
		wordOpt.WriteText(name_vc[i].c_str());
		wordOpt.NewLine();

		wordOpt.SetFont(_T("Times New Roman"),12,0,0); // 小四不加粗
		wordOpt.CreateTable(5,1);

		// 接口详细信息
		wordOpt.WriteCellText(1,1,form_vc[i].c_str());
		wordOpt.WriteCellText(2,1,desc_vc[i].c_str());
		wordOpt.WriteCellText(3,1,param_vc[i].c_str());
		wordOpt.WriteCellText(4,1,return_vc[i].c_str());
		wordOpt.WriteCellText(5,1,example_vc[i].c_str());

		// 设置单元格背景颜色
		wordOpt.SetTableShading(1,1,16777215);
		wordOpt.SetTableShading(2,1,12500670);
		wordOpt.SetTableShading(3,1,16777215);
		wordOpt.SetTableShading(4,1,12500670);
		wordOpt.SetTableShading(5,1,16777215);

		wordOpt.EndLine();
	}

	// 保存并关闭当前Word文档
	regex filename_pattern("(.+?)\\.h");
	vector<string> filename_vc;
	string saved_doc_name;
	if(RegexSearch(filename,filename_pattern,filename_vc,1) == 0){
		saved_doc_name = filename_vc[0] + ".doc";
	}
	wordOpt.SaveAs(saved_doc_name.c_str());           
	wordOpt.Close(); // 关闭当前文档，但不关闭Word程序
	return;
}

void SrcCodeScanner::GenerateMarkdownFile(string filename){
	//////////////////获取生成Markdown文档所需的处理好的数据//////////////////////
	// 定义存放正则匹配结果的容器
	vector<string> class_name_vc;
	vector<string> class_desc_vc;
	vector<string> name_vc;
	vector<string> form_vc;
	vector<string> desc_vc;
	vector<string> param_vc;
	vector<string> return_vc;
	vector<string> example_vc;
	string complete_class;

	// 读取(头)文件内容 -> 执行正则匹配 -> 处理匹配结果 -> 获得需要的处理好的数据
	if (!GetWantedData(filename,complete_class,class_name_vc,class_desc_vc,name_vc,form_vc,desc_vc,param_vc,return_vc,example_vc))
	{
		return; // 若没有获取写入所需的数据，则不再执行生成Markdown文档的操作
	}
	// 对数据进行Markdown语法格式的处理
	string header = "###### GeoBeans平台二次开发库架构及主要接口说明文档";
	string footer = "###### <div style=\"text-align:right\">北京中遥地网信息技术有限公司</div>";
	string format_line = "| :-------- |";
	complete_class = "### **" + complete_class + "**";
	for (size_t i = 0;i < name_vc.size();i++)
	{
		// 处理接口名称
		name_vc[i] = "#### **" + name_vc[i] + "**";

		// 处理接口形式
		form_vc[i] = "| " + form_vc[i] + " |";

		// 处理接口描述
		desc_vc[i] = "| " + desc_vc[i] + " |";

		// 处理接口参数
		ReplaceCharacters(param_vc[i]);
		param_vc[i] = "| " + param_vc[i] + " |";


		// 处理接口返回值
		ReplaceCharacters(return_vc[i]);
		return_vc[i] = "| " + return_vc[i] + " |";


		// 处理接口示例代码
		ReplaceCharacters(example_vc[i]);
		example_vc[i] = "| " + example_vc[i] + " |";

	}

	
	//////////////////写入数据，生成Markdown文档//////////////////
	// 创建并打开Markdown文件
	regex filename_pattern("(.+?)\\.h");
	vector<string> filename_vc;
	string saved_md_name;
	if(RegexSearch(filename,filename_pattern,filename_vc,1) == 0){
		saved_md_name = filename_vc[0] + ".md";
		//AfxMessageBox((CString)saved_md_name.c_str());
	}
	ofstream outfile(saved_md_name,ios::out);

	// 按格式写入数据
	outfile << header << endl;
	outfile << complete_class << endl;
	for (size_t i = 0;i < name_vc.size();i++)
	{
		outfile << name_vc[i] << endl;
		outfile << form_vc[i] << endl;
		outfile << format_line << endl;
		outfile << desc_vc[i] << endl;
		outfile << param_vc[i] << endl;
		outfile << return_vc[i] << endl;
		outfile << example_vc[i] << endl;
	}
	outfile << footer << endl;
	outfile.close();
	return;
}