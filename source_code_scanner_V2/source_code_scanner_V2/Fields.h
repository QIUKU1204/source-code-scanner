// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

//#import "C:\\Program Files\\Microsoft Office\\Office15\\MSWORD.OLB" no_namespace
// Fields 包装类

class Fields : public COleDispatchDriver
{
public:
	Fields(){} // 调用 COleDispatchDriver 默认构造函数
	Fields(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Fields(const Fields& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// Fields 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Locked()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Locked(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	void ToggleShowCodes()
	{
		InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long Update()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_I4, (void*)&result, NULL);
		return result;
	}
	void Unlink()
	{
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void UpdateSource()
	{
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * Type, VARIANT * Text, VARIANT * PreserveFormatting)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Type, Text, PreserveFormatting);
		return result;
	}

	// Fields 属性
public:

};
