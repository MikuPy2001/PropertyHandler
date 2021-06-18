#include <windows.h>

#include <propsys.h>     // Property System APIs and interfaces
#include <propkey.h>     // System PROPERTYKEY definitions
#include <propvarutil.h> // PROPVARIANT and VARIANT helper APIs

#include <new>  // std::nothrow

#include "XmlHelp.h"
#include "Dll.h"
#include "SafeRelease.h"
#include "IPropertyStoreCacheHelp.h"
#include "PropertyHandler.h"

//������������ĵ�����
//������ļ�����ȡ����


class CPropertyHandler :
    public IPropertyStore,// ö�ٺͲ�������ֵ
    public IPropertyStoreCapabilities,// �û��Ƿ������UI�б༭���ԡ�
    public IInitializeWithStream// ʹ������ʼ���������(�����Դ����������ͼ��������Ԥ���������)
{
public:

    CPropertyHandler() : _cRef(1), cache(0){
        OutputDebugString(L"CPropertyHandler::CPropertyHandler()\r\n");
        DllAddRef();//ÿ�γ�ʼ������һ����
    }

    // IUnknown 
    // ���Ӧ���Ǹ��𽫴����õ�com�������͸�ϵͳ
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        OutputDebugString(L"CPropertyHandler::QueryInterface()\r\n");
        //���������ÿһ����Ҫ��¶��ȥ�ĵ�com�ӿ�,���
        static const QITAB qit[] = {
            QITABENT(CPropertyHandler, IPropertyStore),
            QITABENT(CPropertyHandler, IPropertyStoreCapabilities),
            QITABENT(CPropertyHandler, IInitializeWithStream),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() { 
        return InterlockedIncrement(&_cRef);
    }
    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG cRef = InterlockedDecrement(&_cRef);
        if (!cRef) { delete this; }
        return cRef;
    }

    // IPropertyStore
    // �˽ӿڹ�������ö�ٺͲ�������ֵ�ķ�����

    //�˷������ظ��ӵ��ļ����������ļ�����
    IFACEMETHODIMP GetCount(DWORD* pcProps) { *pcProps = 0; return cache ? cache->GetCount(pcProps) : E_UNEXPECTED; }
    //��������������л�ȡ���Լ���
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey) { *pkey = PKEY_Null; return cache ? cache->GetAt(iProp, pkey) : E_UNEXPECTED; }
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar);//�˷��������ض����Ե����ݡ�
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);//�˷�����������ֵ���滻��ɾ������ֵ��
    IFACEMETHODIMP Commit();//�ڽ����˸���֮�󣬴˷���������ġ�

    // IPropertyStoreCapabilities
    // ����һ���������÷���ȷ���û��Ƿ������UI�б༭���ԡ�
    IFACEMETHODIMP IsPropertyWritable(REFPROPERTYKEY key) { return S_FALSE; }//��ֹ�������Ե��޸�

    // IInitializeWithStream
    // ʹ����������ʼ���������(�����Դ����������ͼ��������Ԥ���������)�ķ�����
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode);

private:
    long _cRef;
    IPropertyStoreCache* cache;  // �ڲ�ֵ�����DOM��˳���IPropertyStore����

    // ����ʱִ��
    ~CPropertyHandler() {
        OutputDebugString(L"CPropertyHandler::~CPropertyHandler()\r\n");
        SafeRelease(&cache);
        DllRelease();
    }
};
// ����ʵ�����������͸�ϵͳ,Ȼ������
HRESULT CSongPropertyHandler_CreateInstance(REFIID riid, void** ppv)
{
    OutputDebugString(L"CSongPropertyHandler_CreateInstance()\r\n");
    *ppv = NULL;
    CPropertyHandler* pNew = new(std::nothrow) CPropertyHandler;//������
    HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;//�ж��Ƿ񴴽��ɹ�
    if (SUCCEEDED(hr))
    {
        hr = pNew->QueryInterface(riid, ppv);
        pNew->Release();//����
    }
    return hr;
}


IFACEMETHODIMP CPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar)
{
    OutputDebugString(L"CPropertyHandler::GetValue()\r\n");
    PropVariantInit(pPropVar);
    return cache ? cache->GetValue(key, pPropVar) : E_UNEXPECTED;
}

IFACEMETHODIMP CPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
    OutputDebugString(L"CPropertyHandler::SetValue()\r\n");
    return STG_E_ACCESSDENIED;
}

IFACEMETHODIMP CPropertyHandler::Commit()
{
    OutputDebugString(L"CPropertyHandler::Commit()\r\n");
    return STG_E_ACCESSDENIED;
}
// Map of property keys to the locations of their value(s) in the .recipe XML schema
struct PROPERTYMAP
{
    const PROPERTYKEY* pkey;    // pointer type to enable static declaration
    PCWSTR ·��;
    PCWSTR �ڵ�����;
};

const PROPERTYMAP ȫ��_�б�_PKEY_·��_����[] =
{
    { &PKEY_Title, L"Temp", L"Title" },
    { &PKEY_Author, L"Temp", L"Author" },
};
IFACEMETHODIMP CPropertyHandler::Initialize(IStream* pStream, DWORD grfMode)
{
    OutputDebugString(L"CPropertyHandler::Initialize()\r\n");
    HRESULT hr = E_UNEXPECTED;
    if (!cache) {
        hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&cache));//��ʼ�� �ڴ����Դ洢
        if (hr != S_OK) {
            return hr;
        }
    }

    IXMLDOMDocument* XML = 0;
    hr=XmlHelp(pStream,& XML);
    if (hr != S_OK) {
        return hr;
    }

    for (size_t i = 0; i < 2; i++)
    {
        auto list = ȫ��_�б�_PKEY_·��_����[i];
        BSTR* Xml�ڵ��ı� = 0;
        long XML�ڵ��ı����� = 0;

        hr = XmlHelp_getV(XML, list.·��, list.�ڵ�����, &Xml�ڵ��ı�, &XML�ڵ��ı�����);
        if (hr == S_OK) {
            hr = IPropertyStoreCache_Set_Value(cache, list.pkey, Xml�ڵ��ı�, XML�ڵ��ı�����);
            XmlHelp_getV_Release(Xml�ڵ��ı�, XML�ڵ��ı�����);
        }
    }
    SafeRelease(&XML);

    return hr;
}
