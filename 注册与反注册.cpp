
#include <windows.h>
#include "RegisterExtension.h" // CRegisterExtension
#include "自己的属性处理器.h"


HRESULT RegisterHandler();
//注册
STDAPI DllRegisterServer()
{
    return RegisterHandler();
}

HRESULT UnregisterHandler();
//反注册
STDAPI DllUnregisterServer()
{
    return UnregisterHandler();
}


const WCHAR wchar_文件后缀[] = L".abc";
const WCHAR wchar_文件类型[] = L"MikuPy2001.abc";
const WCHAR wchar_文件描述[] = L"一种自定义文件(.recipe)的处理程序";


//在ProgID上注册proplists以方便取消注册，并最小化与其他可能处理该文件扩展名的应用程序的冲突
//https://docs.microsoft.com/en-us/windows/win32/properties/building-property-handlers-property-lists

//当用户悬停在项目上时，属性将显示在信息尖上。
//仅XP系统有效
const WCHAR c_szRecipeInfoTip[] = L"prop:"
                                            "System.Author"
                                            //"System.ItemType;"
                                            //"System.Author;"
                                            //"System.Rating;"
                                            //"Microsoft.SampleRecipe.Difficulty"
    ;
//属性显示在属性对话框的详细信息选项卡上。这是文件类型支持的属性的完整列表。
const WCHAR c_szRecipeFullDetails[] = L"prop:"
                                            "System.Title;"
                                            "System.Author"

                                            //"System.PropGroup.Description;"
                                            //"System.Title;"
                                            //"System.Author;"
                                            //"System.Comment;"
                                            //"System.Keywords;"
                                            //"System.Rating;"
                                            //"Microsoft.SampleRecipe.Difficulty;"
                                            //"System.PropGroup.FileSystem;"
                                            //"System.ItemNameDisplay;"
                                            //"System.ItemType;"
                                            //"System.ItemFolderPathDisplay;"
                                            //"System.Size;"
                                            //"System.DateCreated;"
                                            //"System.DateModified;"
                                            //"System.DateAccessed;"
                                            //"System.FileAttributes;"
                                            //"System.OfflineAvailability;"
                                            //"System.OfflineStatus;"
                                            //"System.SharedWith;"
                                            //"System.FileOwner;"
                                            //"System.ComputerName"
    ;
//属性显示在预览窗格中。
const WCHAR c_szRecipePreviewDetails[] = L"prop:"
                                               "System.Author"
//                                               "System.DateChanged;"
//                                               "System.Author;"
//                                               "System.Keywords;"
//                                               "Microsoft.SampleRecipe.Difficulty;"
//                                               "System.Rating;"
//                                               "System.Comment;"
//                                               "System.Size;"
//                                               "System.ItemFolderPathDisplay;"
//                                               "System.DateCreated"
    ;
//属性显示在项目缩略图旁边的预览 Pane的标题区域。参赛人数最多为3人。如果属性列表包含的多于最大允许的数字，则会忽略其余条目。
const WCHAR c_szRecipePreviewTitle[] = L"prop:"
                                               "System.Author"
//"System.Title;"
//"System.ItemType"
;
//注册
HRESULT RegisterHandler()
{

    // register the property handler COM object with the system
    // 向系统注册属性处理程序COM对象
    // __uuidof(CRecipePropertyHandler) 拿到 GUID
    CRegisterExtension re(__uuidof(CSongPropertyHandler), HKEY_LOCAL_MACHINE);
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
    CRegisterExtension re(__uuidof(CSongPropertyHandler), HKEY_LOCAL_MACHINE);
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