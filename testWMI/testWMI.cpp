/*********************************************************************************************************
**
** 创   建   人: CSZQ
**
** 描        述: WMI 使用示例
**
*********************************************************************************************************/
/*********************************************************************************************************
	头文件
*********************************************************************************************************/
#include <comdef.h>
#include <Wbemidl.h>

#include <iostream>

#pragma comment(lib, "wbemuuid.lib")

/*********************************************************************************************************
	函数
*********************************************************************************************************/

void testUseWMI(IWbemServices* pSvc) {
	HRESULT hRes = 0;
	IEnumWbemClassObject* pEnumerator = NULL;

	/*
	 * 1、执行查询
	 */
	hRes = pSvc->ExecQuery(
		BSTR(L"WQL"),													// WQL 语言
		BSTR(L"SELECT * FROM Win32_Process"),							// 具体语句
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,			// WBEM_FLAG_FORWARD_ONLY  单向，不能clone和reset，但是速度很快
																		// WBEM_FLAG_RETURN_IMMEDIATELY  半同步调用方法，由WMI管理的类似异步的方法，比异步性能高
		NULL,															// 管他干啥用的，懒得看了，下班了
		&pEnumerator);													// 返回结果

	/*
	 * 2、获取结果
	 */
	IWbemClassObject* pclsObj;
	ULONG uReturn = 0;

	while (pEnumerator) {
		/*
		 * 2.1、从当前位置获取结果
		 */
		hRes = pEnumerator->Next(
			WBEM_INFINITE,												// 超时设置
			1,															// 本次请求的对象个数
			&pclsObj,													// 返回的对象的空间，必须足够大以容纳 uCount 个对象
			&uReturn);													// 实际返回的对象的个数
		
		if (!uReturn) {
			break;
		}

		VARIANT vtProp;

		hRes = pclsObj->Get(											// 获取指定属性的值
			L"Name",													// 属性名称
			0,															// 保留位
			&vtProp,													// 返回的数据保存在这里
			0,															// CIM 的类型，比如 CIM_SINT32 和 CIM_STRING 等，0 自动判断
			0);															// 值的限定，比如是否可以修改
		
		/*
		 * 2.2、使用数据
		 */
		std::wcout << "Process Name : " << vtProp.bstrVal << std::endl;

		/*
		 * 2.3、清理工作
		 */
		VariantClear(&vtProp);

		pclsObj->Release();
		pclsObj = NULL;
	}

	/*
	 * 3、清理工作
	 */
	pEnumerator->Release();
}

/*********************************************************************************************************
	说明：
		WMI 测试
	参数：
		无
	返回值：
		无
*********************************************************************************************************/
void testWMI() {
	HRESULT hRes = 0;
	IWbemLocator* pLoc = NULL;
	IWbemServices* pSvc = 0;

	/*
	 * 1、COM 初始化
	 */
	hRes = CoInitializeEx(
		0,																// 保留
		COINIT_MULTITHREADED);											// 并发环境，多线程安全
	if (hRes != S_OK) {
		std::cout << "Failed to initialize COM library. "
			<< "Error code = 0x" << std::hex << hRes << std::endl;
	}
	
	/*
	 * 2、安全初始化
	 * COM 组件是外部组件，使用外部组件可能需要 用户认证，访问控制，令牌管理等
	 */
	hRes = CoInitializeSecurity(
		NULL,
		-1,      // COM negotiates service                  
		NULL,    // Authentication services
		NULL,    // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,    // authentication
		RPC_C_IMP_LEVEL_IMPERSONATE,  // Impersonation
		NULL,             // Authentication info 
		EOAC_NONE,        // Additional capabilities
		NULL              // Reserved
	);

	/*
	 * 3、创建名称空间接口
	 */
	hRes = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,											// 指定新创建的类对象运行环境
		IID_IWbemLocator,												// 接口的 UID，表示需要使用的接口
		(LPVOID*)&pLoc);												// 返回的接口对象

	if (!pLoc) {
		std::cout << "Failed to connect namespace" << std::endl;
	}

	/*
	 * 4、链接名称空间
	 * 使用当前用户信息
	 */
	hRes = pLoc->ConnectServer(
		BSTR(L"ROOT\\CIMV2"),											// WMI 名称空间
		NULL,															// 用户名称
		NULL,															// 用户密码
		0,																// 语言环境
		NULL,															// 安全标志，0 表示只有连接成功之后才会返回，阻塞
		0,																// 授权使用的域，比如含有域控的情况，需要在哪个域授权
		0,																// 运行上下文对象
		&pSvc);															// 绑定到指定名称空间的对象


	/*
	 * 5、设置连接安全级别
	 */
	hRes = CoSetProxyBlanket(
		pSvc,															// the proxy to set
		RPC_C_AUTHN_WINNT,												// authentication service
		RPC_C_AUTHZ_NONE,												// authorization service
		NULL,															// Server principal name
		RPC_C_AUTHN_LEVEL_CALL,											// authentication level
		RPC_C_IMP_LEVEL_IMPERSONATE,									// impersonation level
		NULL,															// client identity 
		EOAC_NONE														// proxy capabilities     
	);

	/*
	 * 6、使用 WQL 语句进行查询
	 */
	testUseWMI(pSvc);

	/*
	 * 7、收尾
	 */

	pSvc->Release();
	pLoc->Release();
	

	CoUninitialize();

}


/*********************************************************************************************************
	说明：
		主函数
	参数：
		
	返回值：
		
*********************************************************************************************************/
int main()
{
	testWMI();
}
