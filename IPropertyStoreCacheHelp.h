#pragma once

#include <windows.h>
#include <shobjidl.h>

HRESULT IPropertyStoreCache_Set_Value(IPropertyStoreCache* Cache, const PROPERTYKEY* Type, BSTR* 子节点文本数组, long 子节点数量);