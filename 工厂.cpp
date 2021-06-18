#include <windows.h>
#include <new>  // std::nothrow
#include <shlobj.h>
#include <shlwapi.h>
#include <Locale.h>

#include "Dll.h"
#include "PropertyHandler.h"


//这个文件 涉及 COM 编程的一些特殊之处
//我也不甚了解
//大概就是COM组件不能直接使用
//系统总是先获取一个IClassFactory的COM
//再由IClassFactory间接获取目标COM
//但是IClassFactory也是需要额外编写的


typedef HRESULT(*PFNCREATEINSTANCE)(REFIID riid, void** ppvObject);
struct CLASS_OBJECT_INIT
{
    const CLSID* pClsid;//类ID
    PFNCREATEINSTANCE pfnCreate;//类管理函数
};

//允许创建一类对象。
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
            if (clsid == *pClassObjectInits[i].pClsid)//查找对应的类
            {
                IClassFactory* pClassFactory = new (std::nothrow) CClassFactory(pClassObjectInits[i].pfnCreate);//创建基于 类创建方法的工厂类
                hr = pClassFactory ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    hr = pClassFactory->QueryInterface(riid, ppv);//获取ppv
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
    //创建一个未初始化的对象。
    IFACEMETHODIMP CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv)
    {
        OutputDebugString(L"CClassFactory::CreateInstance()\r\n");
        //调用对应的类创建方法创建类给系统
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


// 在这里添加该模块支持的类
const CLASS_OBJECT_INIT c_rgClassObjectInit[] =
{
    { &__uuidof(CPropertyHandler), CSongPropertyHandler_CreateInstance }
};

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv)//系统首先尝试获取一个工厂类
{
    OutputDebugString(L"DllGetClassObject()\r\n");
    return CClassFactory::CreateInstance(clsid, c_rgClassObjectInit, ARRAYSIZE(c_rgClassObjectInit), riid, ppv);
}