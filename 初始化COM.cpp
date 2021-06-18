#include <windows.h>
#include <new>  // std::nothrow
#include <shlobj.h>
#include <shlwapi.h>
#include <Locale.h>
#include "Dll.h"
#include "属性处理器.h"


typedef HRESULT(*PFNCREATEINSTANCE)(REFIID riid, void** ppvObject);
struct CLASS_OBJECT_INIT
{
    const CLSID* pClsid;
    PFNCREATEINSTANCE pfnCreate;
};

// 在这里添加该模块支持的类
const CLASS_OBJECT_INIT c_rgClassObjectInit[] =
{
    { &__uuidof(CRecipePropertyHandler), CRecipePropertyHandler_CreateInstance }
};

//允许创建一类对象。
class CClassFactory : public IClassFactory
{
public:
    static HRESULT CreateInstance(REFCLSID clsid, const CLASS_OBJECT_INIT* pClassObjectInits, size_t cClassObjectInits, REFIID riid, void** ppv)
    {
        *ppv = NULL;
        HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
        for (size_t i = 0; i < cClassObjectInits; i++)
        {
            if (clsid == *pClassObjectInits[i].pClsid)
            {
                IClassFactory* pClassFactory = new (std::nothrow) CClassFactory(pClassObjectInits[i].pfnCreate);
                hr = pClassFactory ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    hr = pClassFactory->QueryInterface(riid, ppv);
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
    //创建一个未初始化的对象。
    IFACEMETHODIMP CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv)
    {
        return punkOuter ? CLASS_E_NOAGGREGATION : _pfnCreate(riid, ppv);
    }

    //锁定在内存中打开的对象应用程序。这使得可以更快地创建实例。
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


STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv)
{
    return CClassFactory::CreateInstance(clsid, c_rgClassObjectInit, ARRAYSIZE(c_rgClassObjectInit), riid, ppv);
}