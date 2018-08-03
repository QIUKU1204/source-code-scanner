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
	// ��һ�ε��÷���ת������ַ�������
	int len = WideCharToMultiByte(CP_OEMCP,NULL,wideText,-1,NULL,0,NULL,FALSE);
	char *multiText; // char*������ʱ���飬��Ϊ��ֵ���м����
	multiText = new char[len];
	// �ڶ��ε��ý� ���ַ��ַ��� ת��Ϊ ���ֽ��ַ���
	WideCharToMultiByte (CP_OEMCP,NULL,wideText,-1,multiText,len,NULL,FALSE);
	str = multiText;// char*ֱ�Ӹ�ֵ��string
	delete []multiText;
}

bool SrcCodeScanner::ReadFileData(string filename,string &data){
	regex ftype_pattern("[^/*""<>|?]+?\\.h");
	if (!regex_match(filename,ftype_pattern))
	{
		AfxMessageBox(_T("�ļ����ʹ�����ѡ��C++ͷ�ļ�(.h)��"));
		return false;
	}
	ifstream in(filename,ios::binary);
	string data_tmp((istreambuf_iterator<char>(in)),
					istreambuf_iterator<char>());
	data = data_tmp;
	// dataΪ�յĴ���
	if (data.empty()) // or data.length == 0
	{
		AfxMessageBox(_T("�Ҳ���ָ�����ļ���"));
		return false;
	}
	return true;
}

bool SrcCodeScanner::RegexSearch(string data,regex pattern,vector<string> &vc,int position){
	// ʹ���� regex_iterator ���ж��ƥ�䣬�������з��ϵ�ƥ����
	auto words_begin = sregex_iterator(data.begin(),data.end(),pattern);
	auto words_end = sregex_iterator();

	// �������ַ������飬��Ҫ��̬�����ڴ�ռ�
	//string * matches;
	int size = 0;
	for (sregex_iterator it = words_begin; it != words_end; ++it)
	{
		vc.push_back(it->format("$" + to_string((long long)position)));
		++size;
	}
	// �� size = 0 ʱ�Ĵ���
	if (size == 0)
	{
		return false;
	}
	return true;

	//matches = new string[size];
	//delete [] matches;
}

void SrcCodeScanner::GenerateWordDoc(string filename,SccWordApi &wordOpt){
	/////////////////////////�ļ���д��������ʽ//////////////////////////

	// ��ȡ�ļ�����(UTF8)
	string filedata;
	if (!ReadFileData(filename,filedata))
	{
		return;
	}

	// ����������ʽ(��Ҫ�Ż�)(ToUTF8)
	regex class_name_pattern(GBKToUTF8("Class ([A-Za-z: ]+) : Begin."));
	regex class_desc_pattern(GBKToUTF8("��˵����(.+)")); // :�룺
	regex name_pattern(GBKToUTF8("����:([A-Za-z~ ]+)"));
	regex form_pattern(GBKToUTF8("(�ӿ���ʽ:.+)"));
	regex desc_pattern(GBKToUTF8("(��������:.+)"));
	regex param_pattern(GBKToUTF8("// ����˵��:((.|\\r|\\n)+?)// ����ֵ"));
	regex return_pattern(GBKToUTF8("// ����ֵ:((.|\\r|\\n)+?)// ʾ��"));
	regex example_pattern(GBKToUTF8("// ʾ��:((.|\\r|\\n)+?)//"));

	// �������ƥ����������
	vector<string> class_name_vc;
	vector<string> class_desc_vc;
	vector<string> name_vc;
	vector<string> form_vc;
	vector<string> desc_vc;
	vector<string> param_vc;
	vector<string> return_vc;
	vector<string> example_vc;

	// ִ��ƥ��(UTF8)
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
		// ������ʾ�������ĸ��ļ�ƥ��ʧ�ܣ�û������Word�ĵ�
		string warning = filename + "ƥ��ʧ�ܣ���鿴��ע�͸�ʽ��";
		AfxMessageBox((CString)warning.c_str());
		return;
	}

	// ��ƥ�������д���(������Ŵ�����)

	string complete_class = GBKToUTF8("1.1.1 ") + class_desc_vc[0] + class_name_vc[0];

	for (unsigned int i = 0;i < name_vc.size();i++)
	{
		// ����ӿ�����(ToUTF8)
		string str_number = std::to_string((long long)(i + 1));
		name_vc[i] = GBKToUTF8("1.1.1 ") + str_number + name_vc[i];

		// ����ӿ�����(ToUTF8)

		// ����ӿڲ���(ToUTF8)
		string str_tmp = param_vc[i]; // ������ͬ�ĵ�����ͬʱ����ͬһ��vector�������˸�ֵ����ʱstring���ͱ���
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end());
		trim(str_tmp);
		param_vc[i] = GBKToUTF8("����: ") + str_tmp;

		// ����ӿڷ���ֵ(ToUTF8)
		str_tmp = return_vc[i];
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'/'),str_tmp.end());
		str_tmp.erase(remove(str_tmp.begin(),str_tmp.end(),'-'),str_tmp.end()); // ������ -1 ��Ϊ 1
		trim(str_tmp);
		return_vc[i] = GBKToUTF8("����ֵ: ") + str_tmp;

		// ����ӿ�ʾ������(ToUTF8)
		str_tmp = example_vc[i];
		trim(str_tmp);
		example_vc[i] = GBKToUTF8("ʾ������: ") + str_tmp;
	}


	////////////////����Word(���ɵ�Word�ĵ�Ĭ��ΪGBK����)//////////////////
	//CsccWordApi wordOpt;
	//wordOpt.CreateApp();
	wordOpt.CreateDocument();

	wordOpt.SetPageSetup(54,54,72,72);
	wordOpt.SetParagraphFormat(3);
	wordOpt.SetHeaderFooter();

	// д����������˵��(ToGBK)
	wordOpt.SetFont(_T("Times New Roman"),14,0,1); //  �ĺżӴ�
	wordOpt.WriteText(UTF8ToGBK(complete_class.c_str()).c_str()); 
	wordOpt.NewLine();

	// ѭ��д��ÿ���ӿں�����ע����Ϣ(ToGBK)
	for (unsigned int i = 0;i < name_vc.size();i++)
	{
		// �ӿں�����
		wordOpt.SetFont(_T("Times New Roman"),12,0,1); // С�ļӴ�
		wordOpt.WriteText(UTF8ToGBK(name_vc[i].c_str()).c_str());
		wordOpt.NewLine();

		wordOpt.SetFont(_T("Times New Roman"),12,0,0); // С�Ĳ��Ӵ�
		wordOpt.CreateTable(5,1);

		// �ӿ���ϸ��Ϣ
		wordOpt.WriteCellText(1,1,UTF8ToGBK(form_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(2,1,UTF8ToGBK(desc_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(3,1,UTF8ToGBK(param_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(4,1,UTF8ToGBK(return_vc[i].c_str()).c_str());
		wordOpt.WriteCellText(5,1,UTF8ToGBK(example_vc[i].c_str()).c_str());

		// ���õ�Ԫ�񱳾���ɫ
		wordOpt.SetTableShading(1,1,16777215);
		wordOpt.SetTableShading(2,1,12500670);
		wordOpt.SetTableShading(3,1,16777215);
		wordOpt.SetTableShading(4,1,12500670);
		wordOpt.SetTableShading(5,1,16777215);

		wordOpt.EndLine();
	}

	// ����Word�ĵ����ر�Word����
	regex filename_pattern("(.+?)\\.h");
	vector<string> filename_vc;
	RegexSearch(filename,filename_pattern,filename_vc,1);
	string saved_doc_name = filename_vc[0] + ".doc";
	// NOTE: string.c_str()���Ը�ֵ��CString��CString���Ը�ֵ��LPCTSTR��
	//       ����string.c_str()������ֱ�Ӹ�ֵ��LPCTSTR��
	//AfxMessageBox((CString)saved_doc_name.c_str()); // CStringW -> LPCTSTR ; 
	wordOpt.SaveAs(saved_doc_name.c_str()); // string.c_str() -> CString;
	wordOpt.Close(); // �رյ�ǰ�ĵ��������ر�Word����
}