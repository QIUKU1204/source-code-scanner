#include "stdafx.h"
#include "src_code_scanner.h"
#include "Word_Api.h"
#include <regex>
#include <iostream>
#include <fstream>
#include <streambuf>
#include <string>
#include <sstream>
#include <vector>
#include <Windows.h>
#include <io.h>
#include <stdlib.h>
#include <direct.h>
#include <time.h>
using namespace std;


SrcCodeScanner::SrcCodeScanner(){
	// 获取当前工作路径   
	_getcwd(working_path, MAX_PATH);
	// 建立、打开日志文件
	log_path = (string)working_path + "\\SrcCodeScanner.log";
	// 在文件末尾追加记录
	logfile.open(log_path,ios::out|ios::app);
}

SrcCodeScanner::~SrcCodeScanner(){
	// 在析构函数中关闭日志文件
	logfile.close();
}

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

	delete[] w_str;
	w_str = NULL;
	delete[] str;
	str = NULL;

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

	delete[] w_sz_GBK;
	w_sz_GBK = NULL;
	delete[] sz_GBK;
	sz_GBK = NULL;

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
	str = multiText; // char*直接赋值给string

	delete []multiText;
	multiText = NULL;
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
				if(strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
					// 递归调用 GetFilesFromFolder
					GetFilesFromFolder(path_begin.assign(folder_path).append("\\").append(fileinfo.name),files_vc,file_count,file_type);  
			}  
			else  // 如果是file_type类型的文件,则加入容器 
			{  
				// 从文件名中提取文件后缀(此方法不能用于从文件路径中提取后缀名)
				string extension;
				string filename = fileinfo.name;
				if (strstr(filename.c_str(), ".") != NULL) // 如果文件拥有后缀(即包含.)
				{
					extension = filename.substr(filename.find_last_of('.'));
				}
				else{                                     // 如果文件没有后缀(即不包含.)
					extension = "";
				}
				
				// 将后缀名与file_type作比较
				if (strcmp(extension.c_str(), file_type.c_str()) == 0) // 必须使file_type的默认值永远不会等于fileinfo.name的后缀ext
				{
					files_vc.push_back(path_begin.assign(folder_path).append("\\").append(fileinfo.name));
					++file_count;
				}
				if (file_type == "\\*") // 必须使file_type的默认值永远不会等于fileinfo.name的后缀ext
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

bool SrcCodeScanner::ReadFileData(string filename,string &filedata){
	// 以二进制读的方式打开文件
	ifstream infile(filename,ios::binary|ios::in);
	if (infile == NULL)
	{
		return false;
	}
	string data_tmp((istreambuf_iterator<char>(infile)),istreambuf_iterator<char>());
	filedata = UTF8ToGBK(data_tmp.c_str()); // 重要的一步: UTF8 -> GBK ;

	// data为空的处理
	if (filedata.empty()) // or data.length == 0
	{
		string warning = filename + " 找不到指定的文件！";
		AfxMessageBox((CString)warning.c_str(),MB_OK|MB_ICONINFORMATION);
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

bool SrcCodeScanner::GetClassBlock(string filename,vector<string> &class_block_vc){
	
	// 获取系统当前时间
	GetLocalTime(&st);
	sprintf_s(str_time,"%d-%d-%d  %d:%d:%d;",st.wYear,st.wMonth,st.wDay,st.wHour,st.wMinute,st.wSecond);

	// 读取文件内容
	string filedata;
	if (!ReadFileData(filename,filedata)) 
	{
		string warning = filename + " 文件读取失败 ";
		logfile << warning << str_time << endl;
		return false; // 文件读取失败
	}
	// 从文件中匹配提取类块
	regex class_block_pattern("(//[ ]*Class[ ]*[A-Za-z]+[:： ]*Begin\\.(.|\\r|\\n)+?)End\\.*"); // 此处不能使用[:：]{1}
	int flag = RegexSearch(filedata,class_block_pattern,class_block_vc,1);
	if (flag != 0)
	{
		// 创建日志文件，将每次匹配的结果(成功或失败)记录在日志文件中
		string warning = filename + " 匹配提取类块失败 ";
		logfile << warning << str_time << endl;
		return false;
	}
	else{
		string warning = filename +  " 匹配提取类块成功 ";
		logfile << warning << str_time << endl;
	}

	return true;
}

bool SrcCodeScanner::GetWantedData(string class_block,int class_number,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc, 
						vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc){

	// 构建正则表达式:兼顾匹配精度与匹配容错率
    // 优化方案:提供接口供用户输入自定义的正则表达式
	regex class_name_pattern("//[ ]*Class[ ]*([A-Za-z]+)[ ]*[:： ]*Begin\\.*"); 
	regex class_desc_pattern("//[ ]*类说明[:： ]*(.+)");             
	regex name_pattern("//[ ]*名称[:： ]*([A-Za-z~ ]+)");             
	regex form_pattern("//[ ]*(接口形式[:：]{1}[ ]*.+)");
	regex desc_pattern("//[ ]*(功能描述[:：]{1}[ ]*.+)");
	regex param_pattern("//[ ]*参数说明[:： ]*((.|\\r|\\n)+?)//[ ]*返回值"); 
	regex return_pattern("//[ ]*返回值[:： ]*((.|\\r|\\n)+?)//[ ]*示例");    
	regex example_pattern("//[ ]*示例[:： ]*((.|\\r|\\n)+?)//");      // 需要优化

	// 执行匹配（成功返回0，失败返回1）
	int flag = RegexSearch(class_block,class_name_pattern,class_name_vc,1);
	flag += RegexSearch(class_block,class_desc_pattern,class_desc_vc,1);
	RegexSearch(class_block,name_pattern,name_vc,1);
	RegexSearch(class_block,form_pattern,form_vc,1);
	RegexSearch(class_block,desc_pattern,desc_vc,1);
	RegexSearch(class_block,param_pattern,param_vc,1);
	RegexSearch(class_block,return_pattern,return_vc,1);
	RegexSearch(class_block,example_pattern,example_vc,1);
	if (flag != 0)    // 若无法匹配到类名和类说明则判定为匹配失败
	{
		return false; // 若无法匹配到类名与类说明则判定为匹配失败
	}

	// 对匹配结果进行处理
	string str_class_number = std::to_string((long long)(class_number));
	complete_class = "1.1." + str_class_number + " " + class_desc_vc[0] + class_name_vc[0];
	for (size_t i = 0;i < name_vc.size();i++)
	{
		// 处理接口名称
		string str_number = std::to_string((long long)(i + 1));
		name_vc[i] = "1.1." + str_class_number + "." + str_number + " " + name_vc[i];

		// 处理接口描述

		// 处理接口描述

		// 处理接口参数
		string str_tmp = param_vc[i]; // 两个不同的迭代器同时操作同一个vector会出错，因此赋值给临时string类型变量
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);                // 清除字符串首尾的空格符、回车符、换行符及制表符
		param_vc[i] = "参数: " + str_tmp;

		// 处理接口返回值
		ReplaceCharacters(return_vc[i],"-1","$$"); // 保护返回值中的-1，防止 -1 变为 1
		str_tmp = return_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end()); 
		trim(str_tmp);
		return_vc[i] = "返回值: " + str_tmp;
		ReplaceCharacters(return_vc[i],"$$","-1"); // 保护返回值中的-1，防止 -1 变为 1

		// 处理接口示例代码
		str_tmp = example_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);
		example_vc[i] = "示例代码: " + str_tmp;
	}
	return true;
}

bool SrcCodeScanner::CheckPathVector(vector<string> &path_vc,HWND hWnd,vector<string> file_extensions){ // 检查存放文件/文件夹路径的vector容器
	if (path_vc.size() == 0) // 检查容器是否为空
	{
		MessageBox(hWnd,_T("请先选择文件/文件夹！"),_T("提示信息"),MB_OK | MB_ICONINFORMATION);
		return false;
	}

	// 检查vc容器中的元素，若为文件夹，则将该文件夹下的所有文件(默认)放入容器中
	// 往后递归迭代的过程中，只会将符合条件的文件放入容器，不存在将子文件夹放入容器的情况
	for (size_t i = 0;i < path_vc.size();i++)
	{
		WIN32_FIND_DATAA FindFileData;
		FindFirstFileA(path_vc[i].c_str(),&FindFileData);
		// 若是文件夹
		if(FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{   
			long file_count = 0;
			string folder_path = path_vc[i];
			path_vc.erase(path_vc.begin() + i);   // 删除文件夹路径
			if (file_extensions.size() == 0)
			{   // 若没有选择扫描文件类型，则默认扫描该文件夹下的所有文件
				GetFilesFromFolder(folder_path,path_vc,file_count);
			} 
			else
			{   // 若选择了扫描文件类型，则依序获取、扫描
				for (size_t j = 0; j < file_extensions.size();j++)
				{
					string file_type = file_extensions[j];
					GetFilesFromFolder(folder_path,path_vc,file_count,file_type);
				}
			}

			if (file_count == 0){ 
				string warning = "文件夹" + folder_path + " 中不包含已选扫描文件类型！";
				MessageBox(hWnd,(CString)warning.c_str(),_T("提示信息"),MB_OK|MB_ICONINFORMATION);
			}
			i--; // 删除vector中的元素后，容器的size将减一，剩下的元素将整体前移一位
		}
		else // 若是文件
		{    
			// 从文件路径中提取文件后缀     
			_splitpath_s(path_vc[i].c_str(),drive,dir,fname,ext);

		    // 若选择了扫描文件类型，则要对文件类型进行检查
			if (file_extensions.size() != 0)  
			{
				int flag = 0;
				string all_extensions;
				for (size_t j = 0;j < file_extensions.size();j++)
				{
					// 当前文件后缀与已选扫描文件类型匹配
					if (strcmp(ext,file_extensions[j].c_str()) == 0)
					{
						flag += 1;
					}
					// 将所有的后缀名连接成一个字符串
					if (file_extensions[j] == "")
					{
						all_extensions = all_extensions + "无后缀类型" + "; ";
					}
					else{
						all_extensions = all_extensions + file_extensions[j] + "; ";
					}
				}
				if (flag == 0) // 当前文件后缀不等于任何已选扫描文件类型
				{
					string warning = path_vc[i] + " 文件类型错误，请选择: " + all_extensions;
					MessageBox(hWnd,(CString)warning.c_str(),_T("提示信息"),MB_OK|MB_ICONINFORMATION);

					// 删除不符合条件的文件路径
					path_vc.erase(path_vc.begin() + i);
					i--; // 删除vector中的元素后，容器的size将减一，剩下的元素将整体前移一位
				}
			} 
		}
		
	}

	if (path_vc.size() == 0) // 再次检查容器是否为空
	{
		return false;
	}

	return true;
}

void SrcCodeScanner::GenerateWordDoc(string filename,vector<string> file_extensions,string encoding,SccWordApi &wordOpt,
								CString header /* =  */,CString footer /* = */ ){
	/////////////////////获取生成Word文档所需的数据//////////////////////////

	// 读取文件 -> 从文件中匹配提取类块
	if (!GetClassBlock(filename,class_block_vc))
	{
		return; // 头文件中的类块匹配提取失败
	}
	

	////////////////////操作Word接口生成GBK编码的Word文档/////////////////////
	try // 用于解决"RPC服务器不可用"的BUG
	{
		wordOpt.CreateDocument(); 
	}
	catch (COleException* e)
	{
		e->Delete();        // 删除Exception 对象，释放其拥有的堆内存，防止内存泄露
		AfxMessageBox(_T("无法启用Word服务，程序需要重新启动。"),MB_OK|MB_ICONINFORMATION);
		PostQuitMessage(0); // 使用exit 退出当前MFC程序将造成内存泄露
	}
	
	wordOpt.SetPageSetup(54,54,72,72);
	wordOpt.SetParagraphFormat(3);
	if (header != "" && footer != "") // 若不为空，则传入页眉页脚
	{
		wordOpt.SetHeaderFooter(encoding.c_str(),header,footer); // 页眉页脚的编码要与正文一致
	} 
	else // 若为空，则使用默认参数
	{
		wordOpt.SetHeaderFooter(encoding.c_str());               // 页眉页脚的编码要与正文一致
	}
	

	// 将每个类的详细信息依次写入文件中
	for (size_t i = 0; i < class_block_vc.size(); i++)
	{
		// 先清空容器，同时赋值字符串为空字符串
		class_name_vc.clear();
		class_desc_vc.clear();
		name_vc.clear();
		form_vc.clear();      
		desc_vc.clear();      
		param_vc.clear();
		return_vc.clear();    
		example_vc.clear();   
		complete_class = "";

		// (对类块)执行正则匹配 -> 处理匹配结果 -> 获得需要的处理好的数据
		if (!GetWantedData(class_block_vc[i],i+1,complete_class,class_name_vc,class_desc_vc,name_vc,form_vc,desc_vc,param_vc,return_vc,example_vc))
		{
			continue;; // 若没有获取写入所需的数据，则不再执行生成Word文档的操作
		}
		
		// 根据选择的编码将数据写入Word文档
		if (encoding == "GBK")
		{
			// 写入类名及类说明
			wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  四号加粗
			wordOpt.WriteText(complete_class.c_str()); 
			wordOpt.NewLine();

			// 循环写入每个接口函数的注释信息
			for (size_t j = 0;j < name_vc.size();j++)
			{
				// 接口函数名
				wordOpt.SetFont(_T("Times New Roman"),12,0,1); // 小四加粗
				wordOpt.WriteText(name_vc[j].c_str());
				wordOpt.NewLine();

				wordOpt.SetFont(_T("Times New Roman"),12,0,0); // 小四不加粗
				wordOpt.CreateTable(5,1);

				// 接口详细信息
				wordOpt.WriteCellText(1,1,form_vc[j].c_str());
				wordOpt.WriteCellText(2,1,desc_vc[j].c_str());
				wordOpt.WriteCellText(3,1,param_vc[j].c_str());
				wordOpt.WriteCellText(4,1,return_vc[j].c_str());
				wordOpt.WriteCellText(5,1,example_vc[j].c_str());

				// 设置单元格背景颜色
				wordOpt.SetTableShading(1,1,16777215);
				wordOpt.SetTableShading(2,1,12500670);
				wordOpt.SetTableShading(3,1,16777215);
				wordOpt.SetTableShading(4,1,12500670);
				wordOpt.SetTableShading(5,1,16777215);

				wordOpt.EndLine();
			}
		} 
		else // UTF-8编码
		{
			// 写入类名及类说明
			wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  四号加粗
			wordOpt.WriteText(GBKToUTF8(complete_class.c_str()).c_str()); 
			wordOpt.NewLine();

			// 循环写入每个接口函数的注释信息
			for (size_t j = 0;j < name_vc.size();j++)
			{
				// 接口函数名
				wordOpt.SetFont(_T("Times New Roman"),12,0,1); // 小四加粗
				wordOpt.WriteText(GBKToUTF8(name_vc[j].c_str()).c_str());
				wordOpt.NewLine();

				wordOpt.SetFont(_T("Times New Roman"),12,0,0); // 小四不加粗
				wordOpt.CreateTable(5,1);

				// 接口详细信息
				wordOpt.WriteCellText(1,1,GBKToUTF8(form_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(2,1,GBKToUTF8(desc_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(3,1,GBKToUTF8(param_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(4,1,GBKToUTF8(return_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(5,1,GBKToUTF8(example_vc[j].c_str()).c_str());

				// 设置单元格背景颜色
				wordOpt.SetTableShading(1,1,16777215);
				wordOpt.SetTableShading(2,1,12500670);
				wordOpt.SetTableShading(3,1,16777215);
				wordOpt.SetTableShading(4,1,12500670);
				wordOpt.SetTableShading(5,1,16777215);

				wordOpt.EndLine();
			}
		}

	}

	// 保存并关闭当前Word文档
	_splitpath_s(filename.c_str(),drive,dir,fname,ext); 
	string saved_doc_name = (string)drive + (string)dir + (string)fname + ".doc";

	wordOpt.SaveAs(saved_doc_name.c_str());           
	wordOpt.Close(); // 关闭当前文档，但不关闭Word程序

	// 主动释放vector 占用的内存
	vector<string>().swap(class_block_vc);
	vector<string>().swap(class_name_vc);
	vector<string>().swap(class_desc_vc);
	vector<string>().swap(name_vc);
	vector<string>().swap(form_vc);
	vector<string>().swap(desc_vc);
	vector<string>().swap(param_vc);
	vector<string>().swap(return_vc);
	vector<string>().swap(example_vc);
	
	return;
}

void SrcCodeScanner::GenerateMarkdownFile(string filename,vector<string> file_extensions,string encoding,
								string header /* =  */, string footer /* = */ ){
	//////////////////获取生成Markdown文档所需的数据//////////////////////
	
	// 读取文件 -> 从文件中匹配提取类块
	if (!GetClassBlock(filename,class_block_vc))
	{
		return; // 头文件中的类块匹配提取失败
	}


	////////////////处理数据，写入数据，生成Markdown文档///////////////////
	// 创建并打开Markdown文件
	_splitpath_s(filename.c_str(),drive,dir,fname,ext); 
	string saved_md_name = (string)drive + (string)dir + (string)fname + ".md";

	ofstream outfile(saved_md_name,ios::out);

	// 定义页眉、页脚与表格样式
	string header_text = "###### " + header;
	string footer_text = "###### <div style=\"text-align:right\">" + footer +"</div>";
	string format_line = "| :-------- |";
	// 若设置了UTF-8编码，则进行转化；
	// 若设置了GBK编码则无需理会，因为程序本身即使用GBK编码；
	if (encoding == "UTF-8")
	{
		header_text = GBKToUTF8(header_text.c_str());
		footer_text = GBKToUTF8(footer_text.c_str());
	} 
	// 写入页眉
	outfile << header_text << endl;

	// 处理数据、写入数据
	for (size_t i = 0; i < class_block_vc.size(); i++)
	{
		// 先清空容器，同时赋值字符串为空字符串
		class_name_vc.clear();
		class_desc_vc.clear();
		name_vc.clear();
		form_vc.clear();      
		desc_vc.clear();      
		param_vc.clear();
		return_vc.clear();    
		example_vc.clear();   
		complete_class = "";

		// (对类块)执行正则匹配 -> 处理匹配结果 -> 获得需要的处理好的数据
		if (!GetWantedData(class_block_vc[i],i+1,complete_class,class_name_vc,class_desc_vc,name_vc,form_vc,desc_vc,param_vc,return_vc,example_vc))
		{
			continue; // 若当前类块处理失败，则继续处理下一个类块
		}

		// 对数据进行Markdown语法格式的处理
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

		if (encoding == "GBK")
		{
			// 按照GBK编码写入数据
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
		} 
		else
		{
			// 按照UTF-8编码写入数据
			outfile << GBKToUTF8(complete_class.c_str()) << endl;
			for (size_t i = 0;i < name_vc.size();i++)
			{
				outfile << GBKToUTF8(name_vc[i].c_str()) << endl;
				outfile << GBKToUTF8(form_vc[i].c_str()) << endl;
				outfile << GBKToUTF8(format_line.c_str()) << endl;
				outfile << GBKToUTF8(desc_vc[i].c_str()) << endl;
				outfile << GBKToUTF8(param_vc[i].c_str()) << endl;
				outfile << GBKToUTF8(return_vc[i].c_str()) << endl;
				outfile << GBKToUTF8(example_vc[i].c_str()) << endl;
			}
		}

	}

	// 写入页脚，关闭Markdown文件
	outfile << footer_text << endl;
	outfile.close();

	// 主动释放vector 占有的内存
	vector<string>().swap(class_block_vc);
	vector<string>().swap(class_name_vc);
	vector<string>().swap(class_desc_vc);
	vector<string>().swap(name_vc);
	vector<string>().swap(form_vc);
	vector<string>().swap(desc_vc);
	vector<string>().swap(param_vc);
	vector<string>().swap(return_vc);
	vector<string>().swap(example_vc);

	return;
}