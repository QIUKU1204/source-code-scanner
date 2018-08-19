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
	// ��ȡ��ǰ����·��   
	_getcwd(working_path, MAX_PATH);
	// ����������־�ļ�
	log_path = (string)working_path + "\\SrcCodeScanner.log";
	// ���ļ�ĩβ׷�Ӽ�¼
	logfile.open(log_path,ios::out|ios::app);
}

SrcCodeScanner::~SrcCodeScanner(){
	// �����������йر���־�ļ�
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
	// ��һ�ε��÷���ת������ַ�������
	int len = WideCharToMultiByte(CP_OEMCP,NULL,wideText,-1,NULL,0,NULL,FALSE);
	char *multiText; // char*������ʱ���飬��Ϊ��ֵ���м����
	multiText = new char[len];
	// �ڶ��ε��ý� ���ַ��ַ��� ת��Ϊ ���ֽ��ַ���
	WideCharToMultiByte (CP_OEMCP,NULL,wideText,-1,multiText,len,NULL,FALSE);
	str = multiText; // char*ֱ�Ӹ�ֵ��string

	delete []multiText;
	multiText = NULL;
}

void SrcCodeScanner::ReplaceCharacters(string &str,string old_value /* =  */,string new_value /* = */ ){
	int sign = str.find(old_value,0); // ���±�Ϊ0��λ�ÿ�ʼƥ��old_value����ƥ���򷵻�����ĸ���±�
	while (sign != string::npos)
	{   
		// ���±�Ϊsign��λ�ÿ�ʼ��ʹ��new_value�滻size���ַ�
		str.replace(sign,new_value.size(),new_value); // �˴�����ǶԵ��滻�����Ե��滻���ܵ������ݶ�ʧ 
		sign = str.find(old_value,sign + 1);
	}
	return;
}

void SrcCodeScanner::GetFilesFromFolder(string folder_path,vector<string>& files_vc,long& file_count,string file_type /* = */ )  
{  
	// �ļ����  
	long hFile = 0;  
	// �ļ���Ϣ  
	struct _finddata_t fileinfo; 

	string path_begin; 
	// ��ʼ����
	if((hFile = _findfirst(path_begin.assign(folder_path).append("\\*").c_str(),&fileinfo)) != -1)  
	{  
		do{  
			// �����Ŀ¼,����֮   
			if((fileinfo.attrib & _A_SUBDIR))  
			{  
				if(strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
					// �ݹ���� GetFilesFromFolder
					GetFilesFromFolder(path_begin.assign(folder_path).append("\\").append(fileinfo.name),files_vc,file_count,file_type);  
			}  
			else  // �����file_type���͵��ļ�,��������� 
			{  
				// ���ļ�������ȡ�ļ���׺(�˷����������ڴ��ļ�·������ȡ��׺��)
				string extension;
				string filename = fileinfo.name;
				if (strstr(filename.c_str(), ".") != NULL) // ����ļ�ӵ�к�׺(������.)
				{
					extension = filename.substr(filename.find_last_of('.'));
				}
				else{                                     // ����ļ�û�к�׺(��������.)
					extension = "";
				}
				
				// ����׺����file_type���Ƚ�
				if (strcmp(extension.c_str(), file_type.c_str()) == 0) // ����ʹfile_type��Ĭ��ֵ��Զ�������fileinfo.name�ĺ�׺ext
				{
					files_vc.push_back(path_begin.assign(folder_path).append("\\").append(fileinfo.name));
					++file_count;
				}
				if (file_type == "\\*") // ����ʹfile_type��Ĭ��ֵ��Զ�������fileinfo.name�ĺ�׺ext
				{
					files_vc.push_back(path_begin.assign(folder_path).append("\\").append(fileinfo.name));
					++file_count;
				}
			}  
		}while(_findnext(hFile, &fileinfo)  == 0); // ��������
		// ��������
		_findclose(hFile);  
	}  
}

bool SrcCodeScanner::ReadFileData(string filename,string &filedata){
	// �Զ����ƶ��ķ�ʽ���ļ�
	ifstream infile(filename,ios::binary|ios::in);
	if (infile == NULL)
	{
		return false;
	}
	string data_tmp((istreambuf_iterator<char>(infile)),istreambuf_iterator<char>());
	filedata = UTF8ToGBK(data_tmp.c_str()); // ��Ҫ��һ��: UTF8 -> GBK ;

	// dataΪ�յĴ���
	if (filedata.empty()) // or data.length == 0
	{
		string warning = filename + " �Ҳ���ָ�����ļ���";
		AfxMessageBox((CString)warning.c_str(),MB_OK|MB_ICONINFORMATION);
		return false;
	}
	return true;
}

int SrcCodeScanner::RegexSearch(string data,regex pattern,vector<string> &vc,int position){
	// ʹ���� regex_iterator ���ж��ƥ�䣬�������з��ϵ�ƥ����
	auto words_begin = sregex_iterator(data.begin(),data.end(),pattern);
	auto words_end = sregex_iterator();

	for (sregex_iterator it = words_begin; it != words_end; ++it)
	{
		vc.push_back(it->format("$" + to_string((long long)position)));
	}

	if (vc.size() == 0) // ƥ��ʧ�ܷ���1
	{
		return 1;
	}

	return 0; // ƥ��ɹ�����0
}

bool SrcCodeScanner::GetClassBlock(string filename,vector<string> &class_block_vc){
	
	// ��ȡϵͳ��ǰʱ��
	GetLocalTime(&st);
	sprintf_s(str_time,"%d-%d-%d  %d:%d:%d;",st.wYear,st.wMonth,st.wDay,st.wHour,st.wMinute,st.wSecond);

	// ��ȡ�ļ�����
	string filedata;
	if (!ReadFileData(filename,filedata)) 
	{
		string warning = filename + " �ļ���ȡʧ�� ";
		logfile << warning << str_time << endl;
		return false; // �ļ���ȡʧ��
	}
	// ���ļ���ƥ����ȡ���
	regex class_block_pattern("(//[ ]*Class[ ]*[A-Za-z]+[:�� ]*Begin\\.(.|\\r|\\n)+?)End\\.*"); // �˴�����ʹ��[:��]{1}
	int flag = RegexSearch(filedata,class_block_pattern,class_block_vc,1);
	if (flag != 0)
	{
		// ������־�ļ�����ÿ��ƥ��Ľ��(�ɹ���ʧ��)��¼����־�ļ���
		string warning = filename + " ƥ����ȡ���ʧ�� ";
		logfile << warning << str_time << endl;
		return false;
	}
	else{
		string warning = filename +  " ƥ����ȡ���ɹ� ";
		logfile << warning << str_time << endl;
	}

	return true;
}

bool SrcCodeScanner::GetWantedData(string class_block,int class_number,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc, 
						vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc){

	// ����������ʽ:���ƥ�侫����ƥ���ݴ���
    // �Ż�����:�ṩ�ӿڹ��û������Զ����������ʽ
	regex class_name_pattern("//[ ]*Class[ ]*([A-Za-z]+)[ ]*[:�� ]*Begin\\.*"); 
	regex class_desc_pattern("//[ ]*��˵��[:�� ]*(.+)");             
	regex name_pattern("//[ ]*����[:�� ]*([A-Za-z~ ]+)");             
	regex form_pattern("//[ ]*(�ӿ���ʽ[:��]{1}[ ]*.+)");
	regex desc_pattern("//[ ]*(��������[:��]{1}[ ]*.+)");
	regex param_pattern("//[ ]*����˵��[:�� ]*((.|\\r|\\n)+?)//[ ]*����ֵ"); 
	regex return_pattern("//[ ]*����ֵ[:�� ]*((.|\\r|\\n)+?)//[ ]*ʾ��");    
	regex example_pattern("//[ ]*ʾ��[:�� ]*((.|\\r|\\n)+?)//");      // ��Ҫ�Ż�

	// ִ��ƥ�䣨�ɹ�����0��ʧ�ܷ���1��
	int flag = RegexSearch(class_block,class_name_pattern,class_name_vc,1);
	flag += RegexSearch(class_block,class_desc_pattern,class_desc_vc,1);
	RegexSearch(class_block,name_pattern,name_vc,1);
	RegexSearch(class_block,form_pattern,form_vc,1);
	RegexSearch(class_block,desc_pattern,desc_vc,1);
	RegexSearch(class_block,param_pattern,param_vc,1);
	RegexSearch(class_block,return_pattern,return_vc,1);
	RegexSearch(class_block,example_pattern,example_vc,1);
	if (flag != 0)    // ���޷�ƥ�䵽��������˵�����ж�Ϊƥ��ʧ��
	{
		return false; // ���޷�ƥ�䵽��������˵�����ж�Ϊƥ��ʧ��
	}

	// ��ƥ�������д���
	string str_class_number = std::to_string((long long)(class_number));
	complete_class = "1.1." + str_class_number + " " + class_desc_vc[0] + class_name_vc[0];
	for (size_t i = 0;i < name_vc.size();i++)
	{
		// ����ӿ�����
		string str_number = std::to_string((long long)(i + 1));
		name_vc[i] = "1.1." + str_class_number + "." + str_number + " " + name_vc[i];

		// ����ӿ�����

		// ����ӿ�����

		// ����ӿڲ���
		string str_tmp = param_vc[i]; // ������ͬ�ĵ�����ͬʱ����ͬһ��vector�������˸�ֵ����ʱstring���ͱ���
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);                // ����ַ�����β�Ŀո�����س��������з����Ʊ��
		param_vc[i] = "����: " + str_tmp;

		// ����ӿڷ���ֵ
		ReplaceCharacters(return_vc[i],"-1","$$"); // ��������ֵ�е�-1����ֹ -1 ��Ϊ 1
		str_tmp = return_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end()); 
		trim(str_tmp);
		return_vc[i] = "����ֵ: " + str_tmp;
		ReplaceCharacters(return_vc[i],"$$","-1"); // ��������ֵ�е�-1����ֹ -1 ��Ϊ 1

		// ����ӿ�ʾ������
		str_tmp = example_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);
		example_vc[i] = "ʾ������: " + str_tmp;
	}
	return true;
}

bool SrcCodeScanner::CheckPathVector(vector<string> &path_vc,HWND hWnd,vector<string> file_extensions){ // ������ļ�/�ļ���·����vector����
	if (path_vc.size() == 0) // ��������Ƿ�Ϊ��
	{
		MessageBox(hWnd,_T("����ѡ���ļ�/�ļ��У�"),_T("��ʾ��Ϣ"),MB_OK | MB_ICONINFORMATION);
		return false;
	}

	// ���vc�����е�Ԫ�أ���Ϊ�ļ��У��򽫸��ļ����µ������ļ�(Ĭ��)����������
	// ����ݹ�����Ĺ����У�ֻ�Ὣ�����������ļ����������������ڽ����ļ��з������������
	for (size_t i = 0;i < path_vc.size();i++)
	{
		WIN32_FIND_DATAA FindFileData;
		FindFirstFileA(path_vc[i].c_str(),&FindFileData);
		// �����ļ���
		if(FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{   
			long file_count = 0;
			string folder_path = path_vc[i];
			path_vc.erase(path_vc.begin() + i);   // ɾ���ļ���·��
			if (file_extensions.size() == 0)
			{   // ��û��ѡ��ɨ���ļ����ͣ���Ĭ��ɨ����ļ����µ������ļ�
				GetFilesFromFolder(folder_path,path_vc,file_count);
			} 
			else
			{   // ��ѡ����ɨ���ļ����ͣ��������ȡ��ɨ��
				for (size_t j = 0; j < file_extensions.size();j++)
				{
					string file_type = file_extensions[j];
					GetFilesFromFolder(folder_path,path_vc,file_count,file_type);
				}
			}

			if (file_count == 0){ 
				string warning = "�ļ���" + folder_path + " �в�������ѡɨ���ļ����ͣ�";
				MessageBox(hWnd,(CString)warning.c_str(),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);
			}
			i--; // ɾ��vector�е�Ԫ�غ�������size����һ��ʣ�µ�Ԫ�ؽ�����ǰ��һλ
		}
		else // �����ļ�
		{    
			// ���ļ�·������ȡ�ļ���׺     
			_splitpath_s(path_vc[i].c_str(),drive,dir,fname,ext);

		    // ��ѡ����ɨ���ļ����ͣ���Ҫ���ļ����ͽ��м��
			if (file_extensions.size() != 0)  
			{
				int flag = 0;
				string all_extensions;
				for (size_t j = 0;j < file_extensions.size();j++)
				{
					// ��ǰ�ļ���׺����ѡɨ���ļ�����ƥ��
					if (strcmp(ext,file_extensions[j].c_str()) == 0)
					{
						flag += 1;
					}
					// �����еĺ�׺�����ӳ�һ���ַ���
					if (file_extensions[j] == "")
					{
						all_extensions = all_extensions + "�޺�׺����" + "; ";
					}
					else{
						all_extensions = all_extensions + file_extensions[j] + "; ";
					}
				}
				if (flag == 0) // ��ǰ�ļ���׺�������κ���ѡɨ���ļ�����
				{
					string warning = path_vc[i] + " �ļ����ʹ�����ѡ��: " + all_extensions;
					MessageBox(hWnd,(CString)warning.c_str(),_T("��ʾ��Ϣ"),MB_OK|MB_ICONINFORMATION);

					// ɾ���������������ļ�·��
					path_vc.erase(path_vc.begin() + i);
					i--; // ɾ��vector�е�Ԫ�غ�������size����һ��ʣ�µ�Ԫ�ؽ�����ǰ��һλ
				}
			} 
		}
		
	}

	if (path_vc.size() == 0) // �ٴμ�������Ƿ�Ϊ��
	{
		return false;
	}

	return true;
}

void SrcCodeScanner::GenerateWordDoc(string filename,vector<string> file_extensions,string encoding,SccWordApi &wordOpt,
								CString header /* =  */,CString footer /* = */ ){
	/////////////////////��ȡ����Word�ĵ����������//////////////////////////

	// ��ȡ�ļ� -> ���ļ���ƥ����ȡ���
	if (!GetClassBlock(filename,class_block_vc))
	{
		return; // ͷ�ļ��е����ƥ����ȡʧ��
	}
	

	////////////////////����Word�ӿ�����GBK�����Word�ĵ�/////////////////////
	try // ���ڽ��"RPC������������"��BUG
	{
		wordOpt.CreateDocument(); 
	}
	catch (COleException* e)
	{
		e->Delete();        // ɾ��Exception �����ͷ���ӵ�еĶ��ڴ棬��ֹ�ڴ�й¶
		AfxMessageBox(_T("�޷�����Word���񣬳�����Ҫ����������"),MB_OK|MB_ICONINFORMATION);
		PostQuitMessage(0); // ʹ��exit �˳���ǰMFC��������ڴ�й¶
	}
	
	wordOpt.SetPageSetup(54,54,72,72);
	wordOpt.SetParagraphFormat(3);
	if (header != "" && footer != "") // ����Ϊ�գ�����ҳüҳ��
	{
		wordOpt.SetHeaderFooter(encoding.c_str(),header,footer); // ҳüҳ�ŵı���Ҫ������һ��
	} 
	else // ��Ϊ�գ���ʹ��Ĭ�ϲ���
	{
		wordOpt.SetHeaderFooter(encoding.c_str());               // ҳüҳ�ŵı���Ҫ������һ��
	}
	

	// ��ÿ�������ϸ��Ϣ����д���ļ���
	for (size_t i = 0; i < class_block_vc.size(); i++)
	{
		// �����������ͬʱ��ֵ�ַ���Ϊ���ַ���
		class_name_vc.clear();
		class_desc_vc.clear();
		name_vc.clear();
		form_vc.clear();      
		desc_vc.clear();      
		param_vc.clear();
		return_vc.clear();    
		example_vc.clear();   
		complete_class = "";

		// (�����)ִ������ƥ�� -> ����ƥ���� -> �����Ҫ�Ĵ���õ�����
		if (!GetWantedData(class_block_vc[i],i+1,complete_class,class_name_vc,class_desc_vc,name_vc,form_vc,desc_vc,param_vc,return_vc,example_vc))
		{
			continue;; // ��û�л�ȡд����������ݣ�����ִ������Word�ĵ��Ĳ���
		}
		
		// ����ѡ��ı��뽫����д��Word�ĵ�
		if (encoding == "GBK")
		{
			// д����������˵��
			wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  �ĺżӴ�
			wordOpt.WriteText(complete_class.c_str()); 
			wordOpt.NewLine();

			// ѭ��д��ÿ���ӿں�����ע����Ϣ
			for (size_t j = 0;j < name_vc.size();j++)
			{
				// �ӿں�����
				wordOpt.SetFont(_T("Times New Roman"),12,0,1); // С�ļӴ�
				wordOpt.WriteText(name_vc[j].c_str());
				wordOpt.NewLine();

				wordOpt.SetFont(_T("Times New Roman"),12,0,0); // С�Ĳ��Ӵ�
				wordOpt.CreateTable(5,1);

				// �ӿ���ϸ��Ϣ
				wordOpt.WriteCellText(1,1,form_vc[j].c_str());
				wordOpt.WriteCellText(2,1,desc_vc[j].c_str());
				wordOpt.WriteCellText(3,1,param_vc[j].c_str());
				wordOpt.WriteCellText(4,1,return_vc[j].c_str());
				wordOpt.WriteCellText(5,1,example_vc[j].c_str());

				// ���õ�Ԫ�񱳾���ɫ
				wordOpt.SetTableShading(1,1,16777215);
				wordOpt.SetTableShading(2,1,12500670);
				wordOpt.SetTableShading(3,1,16777215);
				wordOpt.SetTableShading(4,1,12500670);
				wordOpt.SetTableShading(5,1,16777215);

				wordOpt.EndLine();
			}
		} 
		else // UTF-8����
		{
			// д����������˵��
			wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  �ĺżӴ�
			wordOpt.WriteText(GBKToUTF8(complete_class.c_str()).c_str()); 
			wordOpt.NewLine();

			// ѭ��д��ÿ���ӿں�����ע����Ϣ
			for (size_t j = 0;j < name_vc.size();j++)
			{
				// �ӿں�����
				wordOpt.SetFont(_T("Times New Roman"),12,0,1); // С�ļӴ�
				wordOpt.WriteText(GBKToUTF8(name_vc[j].c_str()).c_str());
				wordOpt.NewLine();

				wordOpt.SetFont(_T("Times New Roman"),12,0,0); // С�Ĳ��Ӵ�
				wordOpt.CreateTable(5,1);

				// �ӿ���ϸ��Ϣ
				wordOpt.WriteCellText(1,1,GBKToUTF8(form_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(2,1,GBKToUTF8(desc_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(3,1,GBKToUTF8(param_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(4,1,GBKToUTF8(return_vc[j].c_str()).c_str());
				wordOpt.WriteCellText(5,1,GBKToUTF8(example_vc[j].c_str()).c_str());

				// ���õ�Ԫ�񱳾���ɫ
				wordOpt.SetTableShading(1,1,16777215);
				wordOpt.SetTableShading(2,1,12500670);
				wordOpt.SetTableShading(3,1,16777215);
				wordOpt.SetTableShading(4,1,12500670);
				wordOpt.SetTableShading(5,1,16777215);

				wordOpt.EndLine();
			}
		}

	}

	// ���沢�رյ�ǰWord�ĵ�
	_splitpath_s(filename.c_str(),drive,dir,fname,ext); 
	string saved_doc_name = (string)drive + (string)dir + (string)fname + ".doc";

	wordOpt.SaveAs(saved_doc_name.c_str());           
	wordOpt.Close(); // �رյ�ǰ�ĵ��������ر�Word����

	// �����ͷ�vector ռ�õ��ڴ�
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
	//////////////////��ȡ����Markdown�ĵ����������//////////////////////
	
	// ��ȡ�ļ� -> ���ļ���ƥ����ȡ���
	if (!GetClassBlock(filename,class_block_vc))
	{
		return; // ͷ�ļ��е����ƥ����ȡʧ��
	}


	////////////////�������ݣ�д�����ݣ�����Markdown�ĵ�///////////////////
	// ��������Markdown�ļ�
	_splitpath_s(filename.c_str(),drive,dir,fname,ext); 
	string saved_md_name = (string)drive + (string)dir + (string)fname + ".md";

	ofstream outfile(saved_md_name,ios::out);

	// ����ҳü��ҳ��������ʽ
	string header_text = "###### " + header;
	string footer_text = "###### <div style=\"text-align:right\">" + footer +"</div>";
	string format_line = "| :-------- |";
	// ��������UTF-8���룬�����ת����
	// ��������GBK������������ᣬ��Ϊ������ʹ��GBK���룻
	if (encoding == "UTF-8")
	{
		header_text = GBKToUTF8(header_text.c_str());
		footer_text = GBKToUTF8(footer_text.c_str());
	} 
	// д��ҳü
	outfile << header_text << endl;

	// �������ݡ�д������
	for (size_t i = 0; i < class_block_vc.size(); i++)
	{
		// �����������ͬʱ��ֵ�ַ���Ϊ���ַ���
		class_name_vc.clear();
		class_desc_vc.clear();
		name_vc.clear();
		form_vc.clear();      
		desc_vc.clear();      
		param_vc.clear();
		return_vc.clear();    
		example_vc.clear();   
		complete_class = "";

		// (�����)ִ������ƥ�� -> ����ƥ���� -> �����Ҫ�Ĵ���õ�����
		if (!GetWantedData(class_block_vc[i],i+1,complete_class,class_name_vc,class_desc_vc,name_vc,form_vc,desc_vc,param_vc,return_vc,example_vc))
		{
			continue; // ����ǰ��鴦��ʧ�ܣ������������һ�����
		}

		// �����ݽ���Markdown�﷨��ʽ�Ĵ���
		complete_class = "### **" + complete_class + "**";
		for (size_t i = 0;i < name_vc.size();i++)
		{
			// ����ӿ�����
			name_vc[i] = "#### **" + name_vc[i] + "**";

			// ����ӿ���ʽ
			form_vc[i] = "| " + form_vc[i] + " |";

			// ����ӿ�����
			desc_vc[i] = "| " + desc_vc[i] + " |";

			// ����ӿڲ���
			ReplaceCharacters(param_vc[i]);
			param_vc[i] = "| " + param_vc[i] + " |";


			// ����ӿڷ���ֵ
			ReplaceCharacters(return_vc[i]);
			return_vc[i] = "| " + return_vc[i] + " |";


			// ����ӿ�ʾ������
			ReplaceCharacters(example_vc[i]);
			example_vc[i] = "| " + example_vc[i] + " |";

		}

		if (encoding == "GBK")
		{
			// ����GBK����д������
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
			// ����UTF-8����д������
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

	// д��ҳ�ţ��ر�Markdown�ļ�
	outfile << footer_text << endl;
	outfile.close();

	// �����ͷ�vector ռ�е��ڴ�
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