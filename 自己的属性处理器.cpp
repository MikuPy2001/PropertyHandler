#include <windows.h>

#include <propsys.h>     // Property System APIs and interfaces
#include <propkey.h>     // System PROPERTYKEY definitions
#include <propvarutil.h> // PROPVARIANT and VARIANT helper APIs

#include <new>  // std::nothrow

#include "自己的属性处理器.h"
#include "XML.h"

#include "Dll.h"

// 等价于 ppt.Release();
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
    public IPropertyStore,// 枚举和操作属性值
    public IPropertyStoreCapabilities,// 用户是否可以在UI中编辑属性。
    public IInitializeWithStream// 使用流初始化处理程序(如属性处理程序、缩略图处理程序或预览处理程序)
{
public:

    CSongPropertyHandler() : _cRef(1), cache(0){
        OutputDebugString(L"CSongPropertyHandler::CSongPropertyHandler()\r\n");
        DllAddRef();//每次初始化都加一次锁
    }

    // IUnknown 
    // 这个应该是负责将创建好的com对象推送给系统
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        OutputDebugString(L"CSongPropertyHandler::QueryInterface()\r\n");
        //在这里加入每一个需要暴露出去的的com接口,大概
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
    // 此接口公开用于枚举和操作属性值的方法。

    //此方法返回附加到文件的属性数的计数。
    IFACEMETHODIMP GetCount(DWORD* pcProps) { *pcProps = 0; return cache ? cache->GetCount(pcProps) : E_UNEXPECTED; }
    //从项的属性数组中获取属性键。
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey) { *pkey = PKEY_Null; return cache ? cache->GetAt(iProp, pkey) : E_UNEXPECTED; }
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar);//此方法检索特定属性的数据。
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);//此方法设置属性值或替换或删除现有值。
    IFACEMETHODIMP Commit();//在进行了更改之后，此方法保存更改。

    // IPropertyStoreCapabilities
    // 公开一个方法，该方法确定用户是否可以在UI中编辑属性。
    IFACEMETHODIMP IsPropertyWritable(REFPROPERTYKEY key) { return S_FALSE; }//禁止所有属性的修改

    // IInitializeWithStream
    // 使用流公开初始化处理程序(如属性处理程序、缩略图处理程序或预览处理程序)的方法。
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode);

private:
    long _cRef;
    IPropertyStoreCache* cache;  // 内部值缓存从DOM后端抽象IPropertyStore操作

    // 结束时执行
    ~CSongPropertyHandler() {
        OutputDebugString(L"CSongPropertyHandler::~CSongPropertyHandler()\r\n");
        SafeRelease(&cache);
        DllRelease();
    }
};
// 负责实例化对象并推送给系统,然后销毁
HRESULT CSongPropertyHandler_CreateInstance(REFIID riid, void** ppv)
{
    OutputDebugString(L"CSongPropertyHandler_CreateInstance()\r\n");
    *ppv = NULL;
    CSongPropertyHandler* pNew = new(std::nothrow) CSongPropertyHandler;//创建类
    HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;//判断是否创建成功
    if (SUCCEEDED(hr))
    {
        hr = pNew->QueryInterface(riid, ppv);
        pNew->Release();//销毁
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
    PCWSTR 路径;
    PCWSTR 节点名称;
};

const PROPERTYMAP 全局_列表_PKEY_路径_名称[] =
{
    { &PKEY_Title, L"Temp", L"Title" },
    { &PKEY_Author, L"Temp", L"Author" },
};
IFACEMETHODIMP CSongPropertyHandler::Initialize(IStream* pStream, DWORD grfMode)
{
    OutputDebugString(L"CSongPropertyHandler::Initialize()\r\n");
    HRESULT hr = E_UNEXPECTED;
    if (!cache) {
        hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&cache));//初始化 内存属性存储
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
        auto list = 全局_列表_PKEY_路径_名称[i];
        BSTR* Xml节点文本 = 0;
        long XML节点文本数量 = 0;

        hr = XmlHelp_getV(XML, list.路径, list.节点名称, &Xml节点文本, &XML节点文本数量);
        if (hr == S_OK) {
            hr = IPropertyStoreCache_Set_Value(cache, list.pkey, Xml节点文本, XML节点文本数量);
            XmlHelp_getV_Release(Xml节点文本, XML节点文本数量);
        }
    }
    SafeRelease(&XML);

    return hr;
}
