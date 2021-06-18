
#include <windows.h>
#include "RegisterExtension.h" // CRegisterExtension
#include "�Լ������Դ�����.h"


HRESULT RegisterHandler();
//ע��
STDAPI DllRegisterServer()
{
    return RegisterHandler();
}

HRESULT UnregisterHandler();
//��ע��
STDAPI DllUnregisterServer()
{
    return UnregisterHandler();
}


const WCHAR wchar_�ļ���׺[] = L".abc";
const WCHAR wchar_�ļ�����[] = L"MikuPy2001.abc";
const WCHAR wchar_�ļ�����[] = L"һ���Զ����ļ�(.recipe)�Ĵ������";


//��ProgID��ע��proplists�Է���ȡ��ע�ᣬ����С�����������ܴ�����ļ���չ����Ӧ�ó���ĳ�ͻ
//https://docs.microsoft.com/en-us/windows/win32/properties/building-property-handlers-property-lists

//���û���ͣ����Ŀ��ʱ�����Խ���ʾ����Ϣ���ϡ�
//��XPϵͳ��Ч
const WCHAR c_szRecipeInfoTip[] = L"prop:"
                                            "System.Author"
                                            //"System.ItemType;"
                                            //"System.Author;"
                                            //"System.Rating;"
                                            //"Microsoft.SampleRecipe.Difficulty"
    ;
//������ʾ�����ԶԻ������ϸ��Ϣѡ��ϡ������ļ�����֧�ֵ����Ե������б�
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
//������ʾ��Ԥ�������С�
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
//������ʾ����Ŀ����ͼ�Աߵ�Ԥ�� Pane�ı������򡣲����������Ϊ3�ˡ���������б�����Ķ��������������֣�������������Ŀ��
const WCHAR c_szRecipePreviewTitle[] = L"prop:"
                                               "System.Author"
//"System.Title;"
//"System.ItemType"
;
//ע��
HRESULT RegisterHandler()
{

    // register the property handler COM object with the system
    // ��ϵͳע�����Դ������COM����
    // __uuidof(CRecipePropertyHandler) �õ� GUID
    CRegisterExtension re(__uuidof(CSongPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.RegisterInProcServer(wchar_�ļ�����, L"Both");
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterInProcServerAttribute(L"ManualSafeSave", TRUE);
        if (SUCCEEDED(hr))
        {
            hr = re.RegisterPropertyHandler(wchar_�ļ���׺);
        }
    }


    if (SUCCEEDED(hr))
    {
        hr = re.RegisterProgIDValue(wchar_�ļ�����, L"FullDetails", c_szRecipeFullDetails);
        if (SUCCEEDED(hr))
        {
            hr = re.RegisterProgIDValue(wchar_�ļ�����, L"PreviewDetails", c_szRecipePreviewDetails);
            if (SUCCEEDED(hr))
            {
                hr = re.RegisterProgIDValue(wchar_�ļ�����, L"PreviewTitle", c_szRecipePreviewTitle);
            }
        }
    }

    //��ProgID���ļ���չ������
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterExtensionWithProgID(wchar_�ļ���׺, wchar_�ļ�����);
    }

    return hr;
}

//��ע��
HRESULT UnregisterHandler()
{
    // Remove the COM object registration.
    CRegisterExtension re(__uuidof(CSongPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();
    if (SUCCEEDED(hr))
    {
        // Unregister the property handler for the file extension.
        hr = re.UnRegisterPropertyHandler(wchar_�ļ���׺);
        if (SUCCEEDED(hr))
        {
            // Remove the whole ProgID since we own all of those settings.
            // Don't try to remove the file extension association since some other application may have overridden it with their own ProgID in the meantime.
            // Leaving the association to a non-existing ProgID is handled gracefully by the Shell.
            // NOTE: If the file extension is unambiguously owned by this application, the association to the ProgID could be safely removed as well,
            //       along with any other association data stored on the file extension itself.
            hr = re.UnRegisterProgID(wchar_�ļ�����, wchar_�ļ���׺);
        }
    }
    return hr;
}