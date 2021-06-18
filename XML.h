#pragma once


#include <windows.h>


HRESULT XmlHelp(IStream* pStream_文件流, IXMLDOMDocument**XML);//你必须 SafeRelease(XML)
HRESULT XmlHelp_getV(IXMLDOMDocument* XML,PCWSTR 父路径, PCWSTR 节点名, BSTR** 子节点文本数组, long* 子节点数量);//你必须 XmlHelp_getV_Release(子节点文本数组,子节点数量)
HRESULT XmlHelp_getV_Release(BSTR* 子节点文本数组, long 子节点数量);


HRESULT IPropertyStoreCache_Set_Value(IPropertyStoreCache* Cache, const PROPERTYKEY* Type, BSTR* 子节点文本数组, long 子节点数量);