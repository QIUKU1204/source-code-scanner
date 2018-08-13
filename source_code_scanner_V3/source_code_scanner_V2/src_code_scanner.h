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
	void trim(string &s);
	string GBKToUTF8(const char* str_GBK);
	string UTF8ToGBK(const char* str_UTF8);
	void WCHARToString(WCHAR *wchar,string &str);
	void ReplaceCharacters(string &str,string old_value = "\r\n",string new_value = "<br>");
	void GetFilesFromFolder(string folder_path,vector<string>& files_vc,long& file_count,string file_type = "\\*");
	bool ReadFileData(string filename,vector<string> file_extensions,string &data);
	int RegexSearch(string data,regex pattern,vector<string> &vc,int position);
	void IterateVector(vector<string> &vc);
	bool GetClassBlock(string filename,vector<string> file_extensions,vector<string> &class_block_vc);
	bool GetWantedData(string class_block,int class_number,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc,
				vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc);
	bool CheckPathVector(vector<string> &path_vc,HWND hWnd,vector<string> file_extensions);
	void GenerateWordDoc(string filename,vector<string> file_extensions,string encoding,SccWordApi &wordOpt,CString header = "",CString footer = "");
	void GenerateMarkdownFile(string filename,vector<string> file_extensions,string encoding,string header = "GeoBeansƽ̨���ο�����ܹ�����Ҫ�ӿ�˵���ĵ�",string footer = "������ң������Ϣ�������޹�˾");

private:
	vector<string> class_block_vc;
	vector<string> class_name_vc;
	vector<string> class_desc_vc;
	vector<string> name_vc;
	vector<string> form_vc;
	vector<string> desc_vc;
	vector<string> param_vc;
	vector<string> return_vc;
	vector<string> example_vc;    // ���ƥ����������
	string complete_class;

	string log_path;              // ��־�ļ�·��
	ofstream logfile;
	char working_path[MAX_PATH];  // ��ǰ����·��
	char drive[_MAX_DRIVE];       // ������
	char dir[_MAX_DIR];           // ·����
	char fname[_MAX_FNAME];       // �ļ���
	char ext[_MAX_EXT];           // ��׺��

	SYSTEMTIME st;
	char str_time[128];           // ϵͳ��ǰʱ��
}; 