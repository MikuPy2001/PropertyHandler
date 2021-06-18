#include "IPropertyStoreCacheHelp.h"
#include <propvarutil.h>

//����ฺ���� IPropertyStoreCache

HRESULT IPropertyStoreCache_Set_Value(IPropertyStoreCache * Cache, const PROPERTYKEY* Type, BSTR* �ӽڵ��ı�����, long �ӽڵ�����) 
{
    OutputDebugString(L"IPropertyStoreCache_Set_Value()\r\n");
    PROPVARIANT ����ֵ = {};
    PropVariantInit(&����ֵ);
    HRESULT hr = InitPropVariantFromStringVector(const_cast<PCWSTR*>(�ӽڵ��ı�����), �ӽڵ�����, &����ֵ);//����
    if (hr == S_OK) {
        hr = PSCoerceToCanonicalValue(*Type, &����ֵ);//ת��
        if (hr == S_OK) {
            hr = Cache->SetValueAndState(*Type, &����ֵ, PSC_NORMAL);//��ֵ
        }
    }
    PropVariantClear(&����ֵ);
    return hr;
}