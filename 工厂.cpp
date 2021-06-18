#include <windows.h>
#include <new>  // std::nothrow
#include <shlobj.h>
#include <shlwapi.h>
#include <Locale.h>

#include "Dll.h"
#include "PropertyHandler.h"


//����ļ� �漰 COM ��̵�һЩ����֮��
//��Ҳ�����˽�
//��ž���COM�������ֱ��ʹ��
//ϵͳ�����Ȼ�ȡһ��IClassFactory��COM
//����IClassFactory��ӻ�ȡĿ��COM
//����IClassFactoryҲ����Ҫ�����д��


typedef HRESULT(*PFNCREATEINSTANCE)(REFIID riid, void** ppvObject);
struct CLASS_OBJECT_INIT
{
    const CLSID* pClsid;//��ID
    PFNCREATEINSTANCE pfnCreate;//�������
};

//������һ�����
class CClassFactory : public IClassFactory
{
public:
    static HRESULT CreateInstance(REFCLSID clsid, const CLASS_OBJECT_INIT* pClassObjectInits, size_t cClassObjectInits, REFIID riid, void** ppv)
    {
        OutputDebugString(L"static CClassFactory::CreateInstance()\r\n");
        *ppv = NULL;
        HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
        for (size_t i = 0; i < cClassObjectInits; i++)
        {
            if (clsid == *pClassObjectInits[i].pClsid)//���Ҷ�Ӧ����
            {
                IClassFactory* pClassFactory = new (std::nothrow) CClassFactory(pClassObjectInits[i].pfnCreate);//�������� �ഴ�������Ĺ�����
                hr = pClassFactory ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    hr = pClassFactory->QueryInterface(riid, ppv);//��ȡppv
                    pClassFactory->Release();
                }
                break; // match found
            }
        }
        return hr;
    }

    CClassFactory(PFNCREATEINSTANCE pfnCreate) : _cRef(1), _pfnCreate(pfnCreate)
    {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        OutputDebugString(L"CClassFactory::QueryInterface()\r\n");
        static const QITAB qit[] =
        {
            QITABENT(CClassFactory, IClassFactory),
            { 0 }
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long cRef = InterlockedDecrement(&_cRef);
        if (cRef == 0)
        {
            delete this;
        }
        return cRef;
    }

    // IClassFactory
    //����һ��δ��ʼ���Ķ���
    IFACEMETHODIMP CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv)
    {
        OutputDebugString(L"CClassFactory::CreateInstance()\r\n");
        //���ö�Ӧ���ഴ�������������ϵͳ
        return punkOuter ? CLASS_E_NOAGGREGATION : _pfnCreate(riid, ppv);
    }

    //�������ڴ��д򿪵Ķ���Ӧ�ó�����ʹ�ÿ��Ը���ش���ʵ����
    IFACEMETHODIMP LockServer(BOOL fLock)
    {
        if (fLock)
        {
            DllAddRef();
        }
        else
        {
            DllRelease();
        }
        return S_OK;
    }

private:
    ~CClassFactory()
    {
        DllRelease();
    }

    long _cRef;
    PFNCREATEINSTANCE _pfnCreate;
};


// ��������Ӹ�ģ��֧�ֵ���
const CLASS_OBJECT_INIT c_rgClassObjectInit[] =
{
    { &__uuidof(CPropertyHandler), CSongPropertyHandler_CreateInstance }
};

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv)//ϵͳ���ȳ��Ի�ȡһ��������
{
    OutputDebugString(L"DllGetClassObject()\r\n");
    return CClassFactory::CreateInstance(clsid, c_rgClassObjectInit, ARRAYSIZE(c_rgClassObjectInit), riid, ppv);
}