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
	*   ȥ���ַ���s ��β�Ŀ��ַ����س��������з����Ʊ��
	*   @param[in] s              ��Ҫ��ϴ���ַ���
	**/
	void trim(string &s);

	/**
	*   ��GBK������ַ���ת��ΪUTF-8����
	*   @param[in] str_GBK        GBK������ַ���
	**/
	string GBKToUTF8(const char* str_GBK);

	/**
	*   ��UTF-8������ַ���ת��ΪGBK����
	*   @param[in] str_UTF8       UTF-8������ַ���
	**/
	string UTF8ToGBK(const char* str_UTF8);

	/**
	*   �����ַ���ת��Ϊ��ͨ���ַ���
	*   @param[in] wchar          ���ַ���
	*   @param[in] str            ��ͨ�ַ���
	**/
	void WCHARToString(WCHAR *wchar,string &str);

	/**
	*   �滻�ַ����е�ָ���ַ������ַ���
	*   @param[in] str            �ַ���
	*   @param[in] old_value      ���滻�����ַ���
	*   @param[in] new_value      �����滻�����ַ���
	**/
	void ReplaceCharacters(string &str,string old_value = "\r\n",string new_value = "<br>");

	/**
	*   ���ļ����л�ȡ���з����������ļ�
	*   @param[in] folder_path     �ļ���·��
	*   @param[in] files_vc        ����ļ�·����vector����
	*   @param[in] file_count      ��������������ļ�����
	*   @param[in] file_type       �ļ�����
	**/
	void GetFilesFromFolder(string folder_path,vector<string>& files_vc,long& file_count,string file_type = "\\*");

	/**
	*   ��ȡ�ļ�����
	*   @param[in] filename         �ļ���·��
	*   @param[in] filedata         �ļ�����
	**/
	bool ReadFileData(string filename,string &filedata);

	/**
	*   ������ִ������ƥ��
	*   @param[in] data             ����
	*   @param[in] pattern          ������ʽ
    *   @param[in] vc               ���ƥ����������
	*   @param[in] position         ƥ�����е����ַ���λ��
	**/
	int RegexSearch(string data,regex pattern,vector<string> &vc,int position);

	/**
	*   ���ļ���������ȡ�������
	*   @param[in] filename         �ļ���·��
	*   @param[in] class_block_vc   ���������������
	**/
	bool GetClassBlock(string filename,vector<string> &class_block_vc);

	/**
	*   ���������������ƥ����ȡע����Ϣ�����д���
	*   @param[in] class_block      �������
	*   @param[in] class_number     ���������
	*   @param[in] complete_class   ��˵�� + ����
	*   @param[in] xxxxxx_vc        ���ע����Ϣ������
	**/
	bool GetWantedData(string class_block,int class_number,string &complete_class,vector<string> &class_name_vc,vector<string> &class_desc_vc,vector<string> &name_vc,
				vector<string> &form_vc,vector<string> &desc_vc,vector<string> &param_vc,vector<string> &return_vc,vector<string> &example_vc);
	
	/**
	*   ����ļ�/�ļ���·������
	*   @param[in] path_vc           �ļ�/�ļ���·������
	*   @param[in] hWnd              ��ǰ���ھ��
	*   @param[in] file_extensions   ��Ҫɨ��/���˵��ļ�����/��׺
	**/
	bool CheckPathVector(vector<string> &path_vc,HWND hWnd,vector<string> file_extensions);

	/**
	*   ������ȡ����ע����Ϣ����Word �ĵ�
	*   @param[in] filename          �ļ���·��
	*   @param[in] file_extensions   ��Ҫɨ��/���˵��ļ�����/��׺
	*   @param[in] encoding          �����ĵ��ı�������
	*   @param[in] wordOpt           Word �ĵ���������
	*   @param[in] header            ҳü�ı�
	*   @param[in] footer            ҳ���ı�
	**/
	void GenerateWordDoc(string filename,vector<string> file_extensions,string encoding,SccWordApi &wordOpt,CString header = "",CString footer = "");
	
	/**
	*   ������ȡ����ע����Ϣ����Markdown �ļ�
	*   @param[in] filename          �ļ���·��
	*   @param[in] file_extensions   ��Ҫɨ��/���˵��ļ�����/��׺
	*   @param[in] encoding          �����ĵ��ı�������
	*   @param[in] header            ҳü�ı�
	*   @param[in] footer            ҳ���ı�
	**/
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