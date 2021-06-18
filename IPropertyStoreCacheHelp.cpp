#include "IPropertyStoreCacheHelp.h"
#include <propvarutil.h>

//这个类负责辅助 IPropertyStoreCache

HRESULT IPropertyStoreCache_Set_Value(IPropertyStoreCache * Cache, const PROPERTYKEY* Type, BSTR* 子节点文本数组, long 子节点数量) 
{
    OutputDebugString(L"IPropertyStoreCache_Set_Value()\r\n");
    PROPVARIANT 返回值 = {};
    PropVariantInit(&返回值);
    HRESULT hr = InitPropVariantFromStringVector(const_cast<PCWSTR*>(子节点文本数组), 子节点数量, &返回值);//设置
    if (hr == S_OK) {
        hr = PSCoerceToCanonicalValue(*Type, &返回值);//转换
        if (hr == S_OK) {
            hr = Cache->SetValueAndState(*Type, &返回值, PSC_NORMAL);//赋值
        }
    }
    PropVariantClear(&返回值);
    return hr;
}