### Note

- Զ�̵��á�����Word����֮ǰ��������Ҫ��ʼ��COM OLE�ӿ�: `AfxOleInit()` ��
- ��ͬһ���̻��̣߳�ֻ�ܵ���һ��`AfxOleInit()`��ʼ��COM���ظ����ý����³������
- (�ظ�)�������ر�Word�������Ϊ���������COM OLE���ʧЧ�������ٴ�����Wordʧ�ܣ�

- `dlg.DoModal()`������һ��ģ̬�Ի���ͬʱ��������߳̽���������ֱ����ģ̬�Ի���رգ�
- ����MFC����Ϣ������:
  - ����Ϣ��������ֵ����֮ǰ�������`afx_msg`��
  - ��Ҫ�ڶԻ������Ե���Ϣѡ����Ϊ����Ϣ��������Ϣ��������

- MessageBox �� AfxMessageBox ������:
  - MessageBox ����ȫ�ֵģ���Dlg.cpp��ʹ�ò���Ҫ���봰�ھ�����������ļ�������Ҫ���ھ����Ϊ������
  - AfxMessageBox ��ȫ�ֵģ����κεط�������Ҫ���ھ����Ϊ������ֱ��ʹ�ü��ɣ�

- CString �� string �Ĺ�ϵ:
  - CStringA ����ֱ�Ӹ�string��ֵ��
  - CStringW ������ֱ�Ӹ�string��ֵ��
  - CStringW -(W2A)-> CStringA -> string ��
  - WCHAR* -(W2A)-> char* -> string ��

- SetDlgItemText �� SetWindowText ����������ͬ��:
  - m_edit_filepath.SetWindowTextW(info); 
  - SetDlgItemTextW(IDC_EDIT,info);

- ���ڱ༭��� EN_CHANGE �� EN_UPDATE �ؼ��¼�:
  - ��༭����ÿ����һ���ַ���EN_CHANGE ����Ӧһ�Σ�
  - ��༭����ÿ����һ���ַ���EN_UPDATE ����Ӧһ�Σ�

- vector������ɾ������ղ���:
  - vector.erase() : ɾ��Ԫ�أ��������տռ䣻ɾ����vector.size��һ��ȫ��Ԫ������ǰ�ƣ�
  - vector.clear() : ���Ԫ�أ��������տռ䣻
  - vector.swap()  : ���Ԫ�أ��������ڴ�ռ䣻

- AfxMessageBox((CString)str.c_str());
- string.c_str()���Ը�ֵ��CString��CString���Ը�ֵ��LPCTSTR������string.c_str()������ֱ�Ӹ�ֵ��LPCTSTR��

- ����������ʽ������Ҫ��:
  - ƥ�侫��Խ�ߣ�ƥ���ݴ���Խ�ͣ�
  - ƥ���ݴ���Խ�ߣ�ƥ�侫��Խ�ͣ�