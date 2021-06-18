#pragma once
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.

#include <windows.h>
#include <new>           // std::nothrow
#include <shobjidl.h>    // IInitializeWithStream, IDestinationStreamFactory
#include <propsys.h>     // Property System APIs and interfaces
#include <propkey.h>     // System PROPERTYKEY definitions
#include <propvarutil.h> // PROPVARIANT and VARIANT helper APIs
#include <msxml6.h>      // DOM interfaces
#include <wincrypt.h>    // CryptBinaryToString, CryptStringToBinary
#include <strsafe.h>     // StringCchPrintf
#include "Dll.h"
#include "XML.h"


// �ȼ��� ppt.Release();
template <class T> void SafeRelease(T** ppT)
{
    if (*ppT)
    {
        OutputDebugString(L"SafeRelease()\r\n");
        (*ppT)->Release();
        *ppT = NULL;
    }
}

__inline BSTR CastToBSTRForInput(PCWSTR psz) { return const_cast<BSTR>(psz); }




HRESULT XmlHelp(IStream* pStream_�ļ���, IXMLDOMDocument** XML)
{
    OutputDebugString(L"XmlHelp()\r\n");
    HRESULT hr = E_UNEXPECTED;
    //��ȡXmlDom��COM����
    hr = CoCreateInstance(
        __uuidof(DOMDocument60),
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(XML)//��ȡ�ɹ����ɴ˱����ش�
    );
    if (SUCCEEDED(hr))
    {
        // �����м���DOM���������
        VARIANT varStream_xml��ȡ = {};
        varStream_xml��ȡ.vt = VT_UNKNOWN;
        varStream_xml��ȡ.punkVal = pStream_�ļ���;
        VARIANT_BOOL vfSuccess_xml��� = VARIANT_FALSE;
        hr = (*XML)->load(varStream_xml��ȡ, &vfSuccess_xml���);//�����ļ�
        if (hr == S_OK && vfSuccess_xml��� == VARIANT_TRUE)
        {
            OutputDebugString(L"XmlHelp() �ļ������ɹ�\r\n");
        }
        else
        {
            OutputDebugString(L"XmlHelp() �ļ�����ʧ��\r\n");
            hr = E_FAIL;
        }

        if (FAILED(hr))
        {
            OutputDebugString(L"XmlHelp() XML �� SafeRelease\r\n");
            SafeRelease(XML);
        }
    }
    return hr;
}

HRESULT XmlHelp_getV(IXMLDOMDocument* XML, PCWSTR ��·��, PCWSTR �ڵ���, BSTR** �ӽڵ��ı�����, long* �ӽڵ�����)
{
    OutputDebugString(L"XmlHelp::getV() ");
    OutputDebugString(��·��);
    OutputDebugString(L" ");
    OutputDebugString(�ڵ���);
    OutputDebugString(L"\r\n");
    // ѡ�����Եĸ��ڵ㲢�������Ե�ֵ
    IXMLDOMNode* ���ڵ� = NULL;
    HRESULT hr = XML->selectSingleNode(CastToBSTRForInput(��·��), &���ڵ�);
    if (hr == S_OK)
    {
        PROPVARIANT propvarValues_xmlֵ = {};

        // select the value nodes
        IXMLDOMNodeList* �ӽڵ��б� = NULL;
        hr = ���ڵ�->selectNodes(CastToBSTRForInput(�ڵ���), &�ӽڵ��б�);

        if (hr == S_OK)
        {
            // ��ȡֵ�ļ���
            *�ӽڵ����� = 0;
            hr = �ӽڵ��б�->get_length(�ӽڵ�����);
            if (SUCCEEDED(hr))
            {
                // ����һ��������������Щֵ
                // wchar_t**
                BSTR* �ӽڵ��ı�����_t = new(std::nothrow) BSTR[*�ӽڵ�����];
                hr = �ӽڵ��ı�����_t ? S_OK : E_OUTOFMEMORY;
                *�ӽڵ��ı����� = �ӽڵ��ı�����_t;
                if (SUCCEEDED(hr))
                {
                    // ��ÿ��ֵ�ڵ���ı����ص�������
                    for (long �ƴ� = 0; �ƴ� < *�ӽڵ�����; �ƴ�++)
                    {
                        if (hr == S_OK)
                        {
                            IXMLDOMNode* �ӽڵ� = NULL;
                            hr = �ӽڵ��б�->get_item(�ƴ�, &�ӽڵ�);
                            if (hr == S_OK)
                            {
                                hr = �ӽڵ�->get_text(&�ӽڵ��ı�����_t[�ƴ�]);
                                OutputDebugString(�ӽڵ��ı�����_t[�ƴ�]);
                                �ӽڵ�->Release();
                            }
                        }
                        else
                        {
                            //ʧ��ʱ��ʣ���Ԫ���ÿ�
                            �ӽڵ��ı�����_t[�ƴ�] = NULL;
                        }
                    }
                }
            }
        }
        ���ڵ�->Release();
    }
    return hr;
}

HRESULT XmlHelp_getV_Release(BSTR* �ӽڵ��ı�����, long �ӽڵ�����)
{
    // ���������ֵ
    for (long iValue = 0; iValue < �ӽڵ�����; iValue++)
    {
        SysFreeString(�ӽڵ��ı�����[iValue]);
    }
    delete[] �ӽڵ��ı�����;

    return S_OK;
}

// PROPVARIANT ����ֵ = {};
// PropVariantClear(&propvarValues_xmlֵ);
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