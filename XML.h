#pragma once


#include <windows.h>


HRESULT XmlHelp(IStream* pStream_�ļ���, IXMLDOMDocument**XML);//����� SafeRelease(XML)
HRESULT XmlHelp_getV(IXMLDOMDocument* XML,PCWSTR ��·��, PCWSTR �ڵ���, BSTR** �ӽڵ��ı�����, long* �ӽڵ�����);//����� XmlHelp_getV_Release(�ӽڵ��ı�����,�ӽڵ�����)
HRESULT XmlHelp_getV_Release(BSTR* �ӽڵ��ı�����, long �ӽڵ�����);


HRESULT IPropertyStoreCache_Set_Value(IPropertyStoreCache* Cache, const PROPERTYKEY* Type, BSTR* �ӽڵ��ı�����, long �ӽڵ�����);