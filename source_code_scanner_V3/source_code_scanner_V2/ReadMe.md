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

- 判断一个路径是文件还是目录(文件夹):
  - 不能将 .xxx 作为文件的依据，因为文件夹同样可以使用 . ；
  - 不能将没有后缀作为目录的依据，因为有的文件可能没有后缀；

- 文件(非文件夹)路径分解:  _splitpath(完整文件路径,磁盘,文件路径,文件名,后缀名);
  - char drive[_MAX_DRIVE];   // 磁盘
  - char dir[_MAX_DIR];       // 路径
  - char fname[_MAX_FNAME];   // 文件名
  - char ext[_MAX_EXT];       // 后缀名
  - _splitpath(filepath,drive,dir,fname,ext); 

- 关于字符串之间的转换: char * <-> string <-> char []
  - string.c_str():生成一个const char *指针，指向一个以'\0'结尾的数组；
  - 使用c_str()后数据将增加一个结束符'\0'；

- 两个BUG:
  - 在一次程序执行过程中，使用Word接口对象wordOpt多次启动、关闭Word程序，将导致无法再次启用Word服务；
  - solution: 一次程序执行过程，只在启动程序的同时创建Word服务，在关闭程序的同时关闭Word；

  - RPC服务器不可用: 在程序执行过程中，手动打开Word文档并关闭，将同时关闭后台运行的由本程序远程调用的Word服务，
    导致无法继续调用Word生成文档；
  - solution: ①使用try...catch 方式捕获并处理wordOpt.AppClose() 引发的COleException 异常； (√)
              ②将后台Word服务与本程序绑定在一起，不允许在外部终止，只能通过AppClose() 关闭；

- 三点新需求(待改进):
  - 创建日志文件，将每次从文件中匹配提取类块的结果(成功或失败)记录在日志文件中； (√)
  - 将Dlg文件中的全局变量封装为Dlg类的数据成员； (√)
  - 将在类的各个函数成员中重复定义使用的变量封装为类的数据成员； (√)
  - 点击退出按钮时，弹窗询问是否退出； (√)

- vector 容器的内存释放机制:
  - vector 的内建中包括内存管理；
  - 当vector 的生命周期/作用域结束时，它的析构函数会把vector 中的元素销毁，同时释放它们所占用的空间，所以vector 一般不用显式释放；
  - 但是，如果vector 中存放的是指针，那么当vector 销毁时，那些指针指向的对象不会被销毁，这些对象拥有的内存也不会被释放；
  - 对于数据量不大的vector，没有必要自己主动释放vector，交给操作系统处理即可；
  - 但是对于存放大量数据的vector，在vector 里面的数据被删除后，主动去释放vector 占有的内存是很有必要的； (√)

- 使用exit 方法强行终止并退出程序，将可能导致内存泄露；为防止内存泄露，应使用PostQuitMessage(0)终止、退出当前程序； (√)


### Update
- 2018.8.15: 完善注释信息；
- 2018.8.16: 调试、修复路径容器path_vc 使用过程中的BUG; 优化警告/提示弹窗的弹出逻辑；
- 2018.8.17: 重新设置主对话框中控件的Tab 键顺序; 使用CDialogEx::UpdateData() 函数完成编辑框控件与关联变量的数据同步；
- 2018.8.18: 优化了正则表达式；
- 2018.8.20: 修改了Word操作函数中对文档页面属性的设置；