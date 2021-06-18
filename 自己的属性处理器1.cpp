#include <windows.h>

#include <propsys.h>     // Property System APIs and interfaces
#include <propkey.h>     // System PROPERTYKEY definitions
#include <propvarutil.h> // PROPVARIANT and VARIANT helper APIs

class CRecipePropertyHandler :
    public IPropertyStore,// 此接口公开用于枚举和操作属性值的方法。
    public IPropertyStoreCapabilities,// 公开一个方法，该方法确定用户是否可以在UI中编辑属性。
    public IInitializeWithStream// 使用流公开初始化处理程序(如属性处理程序、缩略图处理程序或预览处理程序)的方法。
{
public:
    CRecipePropertyHandler() : _cRef(0) { }

    // IUnknown 
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        //在这里加入每一个需要暴露出去的的com接口,大概
        static const QITAB qit[] = {
            QITABENT(CRecipePropertyHandler, IPropertyStore),
            QITABENT(CRecipePropertyHandler, IPropertyStoreCapabilities),
            QITABENT(CRecipePropertyHandler, IInitializeWithStream),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&_cRef); }
    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG cRef = InterlockedDecrement(&_cRef);
        if (!cRef) { delete this; }
        return cRef;
    }

    // IPropertyStore
    // 此接口公开用于枚举和操作属性值的方法。

    //此方法返回附加到文件的属性数的计数。
    IFACEMETHODIMP GetCount(DWORD* pcProps) { *pcProps = 0; return _pCache_IPropertyStoreCache_缓存 ? _pCache_IPropertyStoreCache_缓存->GetCount(pcProps) : E_UNEXPECTED; }
    //从项的属性数组中获取属性键。
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey) { *pkey = PKEY_Null; return _pCache_IPropertyStoreCache_缓存 ? _pCache_IPropertyStoreCache_缓存->GetAt(iProp, pkey) : E_UNEXPECTED; }
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar);//此方法检索特定属性的数据。
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);//此方法设置属性值或替换或删除现有值。
    IFACEMETHODIMP Commit();//在进行了更改之后，此方法保存更改。

    // IPropertyStoreCapabilities
    // 公开一个方法，该方法确定用户是否可以在UI中编辑属性。
    IFACEMETHODIMP IsPropertyWritable(REFPROPERTYKEY key);

    // IInitializeWithStream
    // 使用流公开初始化处理程序(如属性处理程序、缩略图处理程序或预览处理程序)的方法。
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode);

private:
    long _cRef;
    IPropertyStoreCache* _pCache_IPropertyStoreCache_缓存;  // 内部值缓存从DOM后端抽象IPropertyStore操作
    IStream* _pStream_传入Initialize的数据流; // 传入Initialize的数据流，并在提交时保存

    // 结束时执行
    ~CRecipePropertyHandler(){}
};