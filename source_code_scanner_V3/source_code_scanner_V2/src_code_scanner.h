#pragma once
#include <string>
#include <regex>
#include <vector>
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
	void GetFilesFromFolder(string folder_path,vector<string>& files_vc,long& file_count,string file_type = "\\*");
	bool ReadFileData(string filename,string &data);
	bool RegexSearch(string data,regex pattern,vector<string> &vc,int position);
	bool GetWantedData(string filename,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc,
		vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc);
	bool CheckPathVector(vector<string> &path_vc,HWND hWnd);
	void GenerateWordDoc(string filename,SccWordApi &wordOpt);
	void GenerateMarkdownFile(string filename);
private:
};