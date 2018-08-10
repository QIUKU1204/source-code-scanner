### Note

- 远程调用、启动Word程序之前，首先需要初始化COM OLE接口: `AfxOleInit()` ；
- 在同一进程或线程，只能调用一次`AfxOleInit()`初始化COM，重复调用将导致程序出错；
- (重复)创建、关闭Word程序的行为，可能造成COM OLE组件失效，导致再次启动Word失败；

- `dlg.DoModal()`将生成一个模态对话框，同时程序的主线程将被阻塞，直到该模态对话框关闭；
- 关于MFC的消息处理函数:
  - 在消息函数返回值类型之前必须加上`afx_msg`；
  - 需要在对话框属性的消息选项中为该消息设置其消息处理函数；

- MessageBox 与 AfxMessageBox 的区别:
  - MessageBox 不是全局的，在Dlg.cpp中使用不需要传入窗口句柄，在其他文件中则需要窗口句柄作为参数；
  - AfxMessageBox 是全局的，在任何地方都不需要窗口句柄作为参数，直接使用即可；

- CString 与 string 的关系:
  - CStringA 可以直接给string赋值；
  - CStringW 不可以直接给string赋值；
  - CStringW -(W2A)-> CStringA -> string ；
  - WCHAR* -(W2A)-> char* -> string ；

- SetDlgItemText 与 SetWindowText 的作用是相同的:
  - m_edit_filepath.SetWindowTextW(info); 
  - SetDlgItemTextW(IDC_EDIT,info);

- 关于编辑框的 EN_CHANGE 和 EN_UPDATE 控件事件:
  - 向编辑框中每输入一个字符，EN_CHANGE 即响应一次；
  - 向编辑框中每输入一个字符，EN_UPDATE 即响应一次；

- vector容器的删除与清空操作:
  - vector.erase() : 删除元素，但不回收空间；删除后，vector.size减一，全部元素整体前移；
  - vector.clear() : 清除元素，但不回收空间；
  - vector.swap()  : 清除元素，并回收内存空间；

- AfxMessageBox((CString)str.c_str());
- string.c_str()可以赋值给CString，CString可以赋值给LPCTSTR，但是string.c_str()不可以直接赋值给LPCTSTR；

- 关于正则表达式的两个要点:
  - 匹配精度越高，匹配容错率越低；
  - 匹配容错率越高，匹配精度越低；