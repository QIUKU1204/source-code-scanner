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
	bool ReadFileData(string filename,string &data);
	bool RegexSearch(string data,regex pattern,vector<string> &vc,int position);
	void GenerateWordDoc(string filename,SccWordApi &wordOpt);
private:
};