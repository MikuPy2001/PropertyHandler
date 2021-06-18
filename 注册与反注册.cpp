
#include <windows.h>
#include "RegisterExtension.h"
#include "PropertyHandler.h"
#include "��Ϣ.h"

//����ļ���������һЩ����

//ע��
HRESULT RegisterHandler()
{

    // register the property handler COM object with the system
    // ��ϵͳע�����Դ������COM����
    // __uuidof(CRecipePropertyHandler) �õ� GUID
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
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
    CRegisterExtension re(__uuidof(CPropertyHandler), HKEY_LOCAL_MACHINE);
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
//ע��
STDAPI DllRegisterServer()//���� regsvr32.exe PropertyHandler.dll ��ʱ�����
{
    return RegisterHandler();
}

//��ע��
STDAPI DllUnregisterServer()//���� regsvr32.exe /u PropertyHandler.dll ��ʱ�����
{
    return UnregisterHandler();
}