#include <windows.h>

#include <propsys.h>     // Property System APIs and interfaces
#include <propkey.h>     // System PROPERTYKEY definitions
#include <propvarutil.h> // PROPVARIANT and VARIANT helper APIs

#include <new>  // std::nothrow

#include "�Լ������Դ�����.h"
#include "XML.h"

#include "Dll.h"

// �ȼ��� ppt.Release();
template <class T> void SafeRelease(T** ppT)
{
    if (*ppT)
    {
        OutputDebugString(L"SafeRelease()\r\n");
        (*ppT)->Release();
        *ppT = NULL;
    }
}

class CSongPropertyHandler :
    public IPropertyStore,// ö�ٺͲ�������ֵ
    public IPropertyStoreCapabilities,// �û��Ƿ������UI�б༭���ԡ�
    public IInitializeWithStream// ʹ������ʼ���������(�����Դ����������ͼ��������Ԥ���������)
{
public:

    CSongPropertyHandler() : _cRef(1), cache(0){
        OutputDebugString(L"CSongPropertyHandler::CSongPropertyHandler()\r\n");
        DllAddRef();//ÿ�γ�ʼ������һ����
    }

    // IUnknown 
    // ���Ӧ���Ǹ��𽫴����õ�com�������͸�ϵͳ
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        OutputDebugString(L"CSongPropertyHandler::QueryInterface()\r\n");
        //���������ÿһ����Ҫ��¶��ȥ�ĵ�com�ӿ�,���
        static const QITAB qit[] = {
            QITABENT(CSongPropertyHandler, IPropertyStore),
            QITABENT(CSongPropertyHandler, IPropertyStoreCapabilities),
            QITABENT(CSongPropertyHandler, IInitializeWithStream),
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
    ~CSongPropertyHandler() {
        OutputDebugString(L"CSongPropertyHandler::~CSongPropertyHandler()\r\n");
        SafeRelease(&cache);
        DllRelease();
    }
};
// ����ʵ�����������͸�ϵͳ,Ȼ������
HRESULT CSongPropertyHandler_CreateInstance(REFIID riid, void** ppv)
{
    OutputDebugString(L"CSongPropertyHandler_CreateInstance()\r\n");
    *ppv = NULL;
    CSongPropertyHandler* pNew = new(std::nothrow) CSongPropertyHandler;//������
    HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;//�ж��Ƿ񴴽��ɹ�
    if (SUCCEEDED(hr))
    {
        hr = pNew->QueryInterface(riid, ppv);
        pNew->Release();//����
    }
    return hr;
}

#include "jsonxx/json.hpp"
using namespace jsonxx;

IFACEMETHODIMP CSongPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar)
{
    OutputDebugString(L"CSongPropertyHandler::GetValue()\r\n");
    PropVariantInit(pPropVar);
    return cache ? cache->GetValue(key, pPropVar) : E_UNEXPECTED;
}

IFACEMETHODIMP CSongPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
    OutputDebugString(L"CSongPropertyHandler::SetValue()\r\n");
    return STG_E_ACCESSDENIED;
}

IFACEMETHODIMP CSongPropertyHandler::Commit()
{
    OutputDebugString(L"CSongPropertyHandler::Commit()\r\n");
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
IFACEMETHODIMP CSongPropertyHandler::Initialize(IStream* pStream, DWORD grfMode)
{
    OutputDebugString(L"CSongPropertyHandler::Initialize()\r\n");
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
