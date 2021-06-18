
#include <windows.h>
#include "RegisterExtension.h"
#include "PropertyHandler.h"
#include "信息.h"

//这个文件负责配置一些东西

//注册
HRESULT RegisterHandler()
{

    // register the property handler COM object with the system
    // 向系统注册属性处理程序COM对象
    // __uuidof(CRecipePropertyHandler) 拿到 GUID
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.RegisterInProcServer(wchar_文件描述, L"Both");
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterInProcServerAttribute(L"ManualSafeSave", TRUE);
        if (SUCCEEDED(hr))
        {
            hr = re.RegisterPropertyHandler(wchar_文件后缀);
        }
    }


    if (SUCCEEDED(hr))
    {
        hr = re.RegisterProgIDValue(wchar_文件类型, L"FullDetails", c_szRecipeFullDetails);
        if (SUCCEEDED(hr))
        {
            hr = re.RegisterProgIDValue(wchar_文件类型, L"PreviewDetails", c_szRecipePreviewDetails);
            if (SUCCEEDED(hr))
            {
                hr = re.RegisterProgIDValue(wchar_文件类型, L"PreviewTitle", c_szRecipePreviewTitle);
            }
        }
    }

    //将ProgID与文件扩展名关联
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterExtensionWithProgID(wchar_文件后缀, wchar_文件类型);
    }

    return hr;
}

//反注册
HRESULT UnregisterHandler()
{
    // Remove the COM object registration.
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();
    if (SUCCEEDED(hr))
    {
        // Unregister the property handler for the file extension.
        hr = re.UnRegisterPropertyHandler(wchar_文件后缀);
        if (SUCCEEDED(hr))
        {
            // Remove the whole ProgID since we own all of those settings.
            // Don't try to remove the file extension association since some other application may have overridden it with their own ProgID in the meantime.
            // Leaving the association to a non-existing ProgID is handled gracefully by the Shell.
            // NOTE: If the file extension is unambiguously owned by this application, the association to the ProgID could be safely removed as well,
            //       along with any other association data stored on the file extension itself.
            hr = re.UnRegisterProgID(wchar_文件类型, wchar_文件后缀);
        }
    }
    return hr;
}
//注册
STDAPI DllRegisterServer()//运行 regsvr32.exe PropertyHandler.dll 的时候调用
{
    return RegisterHandler();
}

//反注册
STDAPI DllUnregisterServer()//运行 regsvr32.exe /u PropertyHandler.dll 的时候调用
{
    return UnregisterHandler();
}