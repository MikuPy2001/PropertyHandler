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

// MSXML is a dispatch based OM and requires BSTRs as input
// we will cheat and cast C strings to BSTRs as MSXML does not
// utilize the BSTR features of its inputs. this saves the allocations
// that would be necessary otherwise

__inline BSTR CastToBSTRForInput(PCWSTR psz) { return const_cast<BSTR>(psz); }

// �ȼ��� ppt.Release();
template <class T> void SafeRelease(T** ppT)
{
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = NULL;
    }
}

// {1794C9FE-74A9-497f-9C69-B31F03CE7EF9} 100
const PROPERTYKEY PKEY_Microsoft_SampleRecipe_Difficulty = { {0x1794c9fe, 0x74a9, 0x497f, 0x9c, 0x69, 0xb3, 0x1f, 0x03, 0xce, 0x7e, 0xf9}, 100 };

// Map of property keys to the locations of their value(s) in the .recipe XML schema
struct PROPERTYMAP
{
    const PROPERTYKEY* pkey;    // pointer type to enable static declaration
    PCWSTR ·��;
    PCWSTR �ڵ�����;
};

const PROPERTYMAP ȫ��_�б�_PKEY_·��_����[] =
{
    { &PKEY_Title,                             L"Recipe",                L"Title" },
    { &PKEY_Comment,                           L"Recipe",                L"Comments" },
    { &PKEY_Author,                            L"Recipe/Background",     L"Author" },
    { &PKEY_Keywords,                          L"Recipe/RecipeKeywords", L"Keyword" },
    { &PKEY_Microsoft_SampleRecipe_Difficulty, L"Recipe/RecipeInfo",     L"Difficulty" },
};

// Helper functions to opaquely serialize and deserialize PROPVARIANT values to and from string form

HRESULT SerializePropVariantAsString(REFPROPVARIANT propvar, PWSTR* pszOut);
HRESULT DeserializePropVariantFromString(PCWSTR pszIn, PROPVARIANT* ppropvar);

// DLL lifetime management functions

void DllAddRef();
void DllRelease();

class CRecipePropertyHandler :
    public IPropertyStore,
    public IPropertyStoreCapabilities,
    public IInitializeWithStream
{
public:
    CRecipePropertyHandler() : _cRef(1), _���Ա_�ļ���(NULL), _grfMode_��Initialize��STGMģʽ(0), _���Ա_XmlDom_recipe�ļ�(NULL), _���Ա_IPropertyStoreCache_����(NULL)
    {
        DllAddRef();//ÿ�γ�ʼ������һ����
    }

    // IUnknown 
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        static const QITAB qit[] = {
            QITABENT(CRecipePropertyHandler, IPropertyStore),
            QITABENT(CRecipePropertyHandler, IPropertyStoreCapabilities),
            QITABENT(CRecipePropertyHandler, IInitializeWithStream),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG cRef = InterlockedDecrement(&_cRef);
        if (!cRef)
        {
            delete this;
        }
        return cRef;
    }

    // IPropertyStore
    // �˽ӿڹ�������ö�ٺͲ�������ֵ�ķ�����
    IFACEMETHODIMP GetCount(DWORD* pcProps);//�˷������ظ��ӵ��ļ����������ļ�����
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey);//��������������л�ȡ���Լ���
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar);//�˷��������ض����Ե����ݡ�
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);//�˷�����������ֵ���滻��ɾ������ֵ��
    IFACEMETHODIMP Commit();//�ڽ����˸���֮�󣬴˷���������ġ�

    // IPropertyStoreCapabilities
    // ����һ���������÷���ȷ���û��Ƿ������UI�б༭���ԡ�
    IFACEMETHODIMP IsPropertyWritable(REFPROPERTYKEY key);

    // IInitializeWithStream
    // ʹ����������ʼ���������(�����Դ����������ͼ��������Ԥ���������)�ķ�����
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode);

private:

    //ÿ��class�����ٶ�
    ~CRecipePropertyHandler()
    {
        SafeRelease(&_���Ա_�ļ���);
        SafeRelease(&_���Ա_XmlDom_recipe�ļ�);
        SafeRelease(&_���Ա_IPropertyStoreCache_����);
        DllRelease();
    }

    // helpers to load data from the DOM
    HRESULT _LoadCacheFromDom();
    HRESULT _LoadPropertyValues(IXMLDOMNode* pNodeParent, PCWSTR pszNodeValues, PROPVARIANT* ppropvar);
    HRESULT _LoadProperty(const PROPERTYMAP& map);
    HRESULT _LoadExtendedProperties();
    HRESULT _LoadSearchContent();

    // helpers to save data to the DOM
    HRESULT _SaveCacheToDom();
    HRESULT _SavePropertyValues(IXMLDOMNode* pNodeParent, PCWSTR pszNodeValues, REFPROPVARIANT propvar);
    HRESULT _SaveProperty(REFPROPVARIANT propvar, const PROPERTYMAP& map);
    HRESULT _SaveExtendedProperty(REFPROPERTYKEY key, REFPROPVARIANT propvar);
    HRESULT _EnsureChildNodeExists(IXMLDOMNode* pNodeParent, PCWSTR pszName, PCWSTR pszXPath, IXMLDOMNode** ppNodeChild);

    long _cRef;
    IStream* _���Ա_�ļ���; // ����Initialize���������������ύʱ����
    DWORD                _grfMode_��Initialize��STGMģʽ; // ���ݸ�Initialize��STGMģʽ
    IXMLDOMDocument* _���Ա_XmlDom_recipe�ļ�; // ��ʾ.recipe�ļ���DOM����
    IPropertyStoreCache* _���Ա_IPropertyStoreCache_����;  // �ڲ�ֵ�����DOM��˳���IPropertyStore����
};
// ����ʵ�����������͸�ϵͳ,Ȼ������
HRESULT CRecipePropertyHandler_CreateInstance(REFIID riid, void** ppv)
{
    *ppv = NULL;
    CRecipePropertyHandler* pNew = new(std::nothrow) CRecipePropertyHandler;
    HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pNew->QueryInterface(riid, ppv);
        pNew->Release();
    }
    return hr;
}
//���ʷ���ֱ��ת�����ڲ�ֵ����
HRESULT CRecipePropertyHandler::GetCount(DWORD* pcProps){    *pcProps = 0;    return _���Ա_IPropertyStoreCache_���� ? _���Ա_IPropertyStoreCache_����->GetCount(pcProps) : E_UNEXPECTED;}

HRESULT CRecipePropertyHandler::GetAt(DWORD iProp, PROPERTYKEY* pkey)
{
    *pkey = PKEY_Null;
    return _���Ա_IPropertyStoreCache_���� ? _���Ա_IPropertyStoreCache_����->GetAt(iProp, pkey) : E_UNEXPECTED;
}

HRESULT CRecipePropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar)
{
    PropVariantInit(pPropVar);
    return _���Ա_IPropertyStoreCache_���� ? _���Ա_IPropertyStoreCache_����->GetValue(key, pPropVar) : E_UNEXPECTED;
}

// SetValueֻ�Ǹ����ڲ�ֵ����
HRESULT CRecipePropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
    HRESULT hr = E_UNEXPECTED;
    if (_���Ա_IPropertyStoreCache_����)
    {
        // check grfMode to ensure writes are allowed
        hr = STG_E_ACCESSDENIED;
        if ((_grfMode_��Initialize��STGMģʽ & STGM_READWRITE) &&
            (key != PKEY_Search_Contents))  // this property is read-only
        {
            hr = _���Ա_IPropertyStoreCache_����->SetValueAndState(key, &propVar, PSC_DIRTY);
        }
    }

    return hr;
}

// Commit���ڲ�ֵ����д�ش��ݸ�Initialize����
HRESULT CRecipePropertyHandler::Commit()
{
    HRESULT hr = E_UNEXPECTED;
    if (_���Ա_IPropertyStoreCache_����)
    {
        // ���grfMode��ȷ������д
        hr = STG_E_ACCESSDENIED;
        if (_grfMode_��Initialize��STGMģʽ & STGM_READWRITE)
        {
            // ���ڲ�ֵ���汣�浽XML DOM����
            hr = _SaveCacheToDom();
            if (SUCCEEDED(hr))
            {
                // ���������
                LARGE_INTEGER liZero = {};
                hr = _���Ա_�ļ���->Seek(liZero, STREAM_SEEK_SET, NULL);
                if (SUCCEEDED(hr))
                {
                    // ��ȡ��ʱĿ���������ֶ���ȫ����
                    IDestinationStreamFactory* pSafeCommit;
                    hr = _���Ա_�ļ���->QueryInterface(&pSafeCommit);
                    if (SUCCEEDED(hr))
                    {
                        IStream* pStreamCommit;
                        hr = pSafeCommit->GetDestinationStream(&pStreamCommit);
                        //��ȡһ����������������Ҫ���Ƶ��ļ����°汾��
                        if (SUCCEEDED(hr))
                        {
                            // ��XMLд����ʱ�����ύ��
                            VARIANT varStream = {};
                            varStream.vt = VT_UNKNOWN;
                            varStream.punkVal = pStreamCommit;
                            hr = _���Ա_XmlDom_recipe�ļ�->save(varStream);
                            if (SUCCEEDED(hr))
                            {
                                hr = pStreamCommit->Commit(STGC_DEFAULT);
                                if (SUCCEEDED(hr))
                                {
                                    // �ύ��ʵ�������
                                    _���Ա_�ļ���->Commit(STGC_DEFAULT);
                                }
                            }

                            pStreamCommit->Release();
                        }

                        pSafeCommit->Release();
                    }
                }
            }
        }
    }

    return hr;
}

//ָʾ�û��Ƿ�Ӧ���ܹ��༭�������Լ���ֵ
HRESULT CRecipePropertyHandler::IsPropertyWritable(REFPROPERTYKEY key)
{
    // System.Search.Contents ��Ψһ��֧��д�������
    return (key == PKEY_Search_Contents) ? S_FALSE : S_OK;
}

// Initializeʹ������ָ��������������ڲ�ֵ����
HRESULT CRecipePropertyHandler::Initialize(IStream* pStream_�ļ���, DWORD grfMode)
{
    HRESULT hr = E_UNEXPECTED;
    if (!_���Ա_�ļ���)//��ֹ��ʼ������?
    {
        //��ȡXmlDom��COM����
        hr = CoCreateInstance(
            __uuidof(DOMDocument60), 
            NULL, 
            CLSCTX_INPROC_SERVER, 
            IID_PPV_ARGS(&_���Ա_XmlDom_recipe�ļ�)//��ȡ�ɹ����ɴ˱����ش�
        );
        if (SUCCEEDED(hr))
        {
            // �����м���DOM���������
            VARIANT varStream_xml��ȡ = {};
            varStream_xml��ȡ.vt = VT_UNKNOWN;
            varStream_xml��ȡ.punkVal = pStream_�ļ���;
            VARIANT_BOOL vfSuccess_xml��� = VARIANT_FALSE;
            hr = _���Ա_XmlDom_recipe�ļ�->load(varStream_xml��ȡ, &vfSuccess_xml���);//�����ļ�
            if (hr == S_OK && vfSuccess_xml��� == VARIANT_TRUE)
            {
                // ��DOM��������ڲ�ֵ����
                hr = _LoadCacheFromDom();
                if (SUCCEEDED(hr))
                {
                    // ���������grfMode������
                    hr = pStream_�ļ���->QueryInterface(IID_PPV_ARGS(&_���Ա_�ļ���));
                    if (SUCCEEDED(hr))
                    {
                        _grfMode_��Initialize��STGMģʽ = grfMode;
                    }
                }
            }
            else
            {
                hr = E_FAIL;
            }

            if (FAILED(hr))
            {
                SafeRelease(&_���Ա_XmlDom_recipe�ļ�);
            }
        }
    }

    return hr;
}

//���ڲ�DOM��������ڲ�ֵ����
//���� _���Ա_XmlDom_recipe�ļ�
//��ŵ� cache
HRESULT CRecipePropertyHandler::_LoadCacheFromDom()
{
    HRESULT hr = S_OK;

    if (!_���Ա_IPropertyStoreCache_����)
    {
        // �����ڲ�ֵ����
        hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&_���Ա_IPropertyStoreCache_����));//��ʼ�� �ڴ����Դ洢
        if (SUCCEEDED(hr))
        {
            // ֱ�Ӵ�XML��䱾������
            for (UINT i = 0; i < ARRAYSIZE(ȫ��_�б�_PKEY_·��_����); ++i)
            {
                _LoadProperty(ȫ��_�б�_PKEY_·��_����[i]);
            }

            // ������չ���Ժ���������
            _LoadExtendedProperties();
            _LoadSearchContent();
        }
    }

    return hr;
}

//�Ӹ����ĸ��ڵ����ָ����ֵ���������Ǵ���PROPVARIANT
HRESULT CRecipePropertyHandler::_LoadPropertyValues(IXMLDOMNode* ���ڵ�,
    PCWSTR �ӽڵ�����,
    PROPVARIANT* ����ֵ)
{
    // ��ʼ��
    PropVariantInit(����ֵ);

    // select the value nodes
    IXMLDOMNodeList* �ӽڵ��б� = NULL;
    HRESULT hr = ���ڵ�->selectNodes(CastToBSTRForInput(�ӽڵ�����), &�ӽڵ��б�);
    if (hr == S_OK)
    {
        // ��ȡֵ�ļ���
        long �ӽڵ����� = 0;
        hr = �ӽڵ��б�->get_length(&�ӽڵ�����);
        if (SUCCEEDED(hr))
        {
            // ����һ��������������Щֵ
            // wchar_t**
            BSTR* �ӽڵ��ı����� = new(std::nothrow) BSTR[�ӽڵ�����];
            hr = �ӽڵ��ı����� ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                // ��ÿ��ֵ�ڵ���ı����ص�������
                for (long �ƴ� = 0; �ƴ� < �ӽڵ�����; �ƴ�++)
                {
                    if (hr == S_OK)
                    {
                        IXMLDOMNode* �ӽڵ� = NULL;
                        hr = �ӽڵ��б�->get_item(�ƴ�, &�ӽڵ�);
                        if (hr == S_OK)
                        {
                            hr = �ӽڵ�->get_text(&�ӽڵ��ı�����[�ƴ�]);
                            �ӽڵ�->Release();
                        }
                    }
                    else
                    {
                        //ʧ��ʱ��ʣ���Ԫ���ÿ�
                        �ӽڵ��ı�����[�ƴ�] = NULL;
                    }
                }

                if (hr == S_OK)
                {
                    // ��ֵ���б�����PROPVARIANT��
                    hr = InitPropVariantFromStringVector(const_cast<PCWSTR*>(�ӽڵ��ı�����), �ӽڵ�����, ����ֵ);
                }

                // ���������ֵ
                for (long iValue = 0; iValue < �ӽڵ�����; iValue++)
                {
                    SysFreeString(�ӽڵ��ı�����[iValue]);
                }
                delete[] �ӽڵ��ı�����;
            }
        }
    }

    return hr;
}

//���ظ���ӳ����ָ�������Ե����ݵ��ڲ�ֵ������
HRESULT CRecipePropertyHandler::_LoadProperty(const PROPERTYMAP& PKEY_·��_����)
{
    // ѡ�����Եĸ��ڵ㲢�������Ե�ֵ
    IXMLDOMNode* ���ڵ� = NULL;
    HRESULT hr = _���Ա_XmlDom_recipe�ļ�->selectSingleNode(CastToBSTRForInput(PKEY_·��_����.·��), &���ڵ�);
    if (hr == S_OK)
    {
        PROPVARIANT propvarValues_xmlֵ = {};
        hr = _LoadPropertyValues(���ڵ�, PKEY_·��_����.�ڵ�����, &propvarValues_xmlֵ);
        if (hr == S_OK)
        {
            // ǿ�ƽ�ֵת��Ϊ���Լ����ʵ�����
            hr = PSCoerceToCanonicalValue(*PKEY_·��_����.pkey, &propvarValues_xmlֵ);
            if (SUCCEEDED(hr))
            {
                // ������ص�ֵ
                hr = _���Ա_IPropertyStoreCache_����->SetValueAndState(*PKEY_·��_����.pkey, &propvarValues_xmlֵ, PSC_NORMAL);
                // PSC_NORMAL �Ʋ�δ���ı䡣
                // PSC_NOTINSOURCE	��ʼ�����Դ��������ļ��������������Բ����ڡ�
                // PSC_DIRTY	�����ѱ����ģ�����δ�ύ���ļ�������
                // PSC_READONLY
            }
            PropVariantClear(&propvarValues_xmlֵ);
        }
        ���ڵ�->Release();
    }

    return hr;
}

//�����κ��ⲿ���Ե�����(������Щû����ʽӳ�䵽XMLģʽ������)���ڲ�ֵ����
HRESULT CRecipePropertyHandler::_LoadExtendedProperties()
{
    // select the list of extended property nodes
    IXMLDOMNodeList* pList = NULL;
    HRESULT hr = _���Ա_XmlDom_recipe�ļ�->selectNodes(CastToBSTRForInput(L"Recipe/ExtendedProperties/Property"), &pList);
    if (hr == S_OK)
    {
        long cElems = 0;
        hr = pList->get_length(&cElems);
        if (hr == S_OK)
        {
            // iterate over the list and cache each value
            for (long iElem = 0; iElem < cElems; ++iElem)
            {
                IXMLDOMNode* pNode = NULL;
                hr = pList->get_item(iElem, &pNode);
                if (hr == S_OK)
                {
                    IXMLDOMElement* pElement = NULL;
                    hr = pNode->QueryInterface(IID_PPV_ARGS(&pElement));
                    if (SUCCEEDED(hr))
                    {
                        // get the name of the property and convert it to a PROPERTYKEY
                        VARIANT varPropKey = {};
                        hr = pElement->getAttribute(CastToBSTRForInput(L"Key"), &varPropKey);
                        if (hr == S_OK)
                        {
                            PROPERTYKEY key;
                            hr = PSPropertyKeyFromString(varPropKey.bstrVal, &key);
                            if (SUCCEEDED(hr))
                            {
                                // get the encoded value and deserialize it into a PROPVARIANT
                                VARIANT varEncodedValue = {};
                                hr = pElement->getAttribute(CastToBSTRForInput(L"EncodedValue"), &varEncodedValue);
                                if (hr == S_OK)
                                {
                                    PROPVARIANT propvarValue = {};
                                    hr = DeserializePropVariantFromString(varEncodedValue.bstrVal, &propvarValue);
                                    if (SUCCEEDED(hr))
                                    {
                                        // cache the value loaded
                                        hr = _���Ա_IPropertyStoreCache_����->SetValueAndState(key, &propvarValue, PSC_NORMAL);
                                        PropVariantClear(&propvarValue);
                                    }

                                    VariantClear(&varEncodedValue);
                                }
                            }

                            VariantClear(&varPropKey);
                        }

                        pElement->Release();
                    }
                    pNode->Release();
                }
            }
        }
        pList->Release();
    }

    return hr;
}

//���System.Search���ڲ�ֵ�����е���������
HRESULT CRecipePropertyHandler::_LoadSearchContent()
{
    // XSLT to generate a space-delimited list of Items, Steps, Yield, Difficulty, and Keywords
    BSTR bstrContentXSLT = SysAllocString(L"<xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\">\n"
        L"<xsl:output method=\"text\" version=\"1.0\" encoding=\"UTF-8\" indent=\"no\"/>\n"
        L"  <xsl:template match=\"/\">\n"
        L"    <xsl:apply-templates select=\"Recipe/Ingredients/Item\"/>\n"
        L"    <xsl:apply-templates select=\"Recipe/Directions/Step\"/>\n"
        L"    <xsl:apply-templates select=\"Recipe/RecipeInfo/Yield\"/>\n"
        L"    <xsl:apply-templates select=\"Recipe/RecipeInfo/Difficulty\"/>\n"
        L"    <xsl:apply-templates select=\"Recipe/RecipeKeywords/Keyword\"/>\n"
        L"  </xsl:template>\n"
        L"  <xsl:template match=\"*\">\n"
        L"    <xsl:value-of select=\".\"/>\n"
        L"    <xsl:text> </xsl:text>\n"
        L"  </xsl:template>\n"
        L"</xsl:stylesheet>");
    HRESULT hr = bstrContentXSLT ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // create a DOM object to hold the XSLT
        IXMLDOMDocument* pContentXSLT = NULL;
        hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&pContentXSLT));
        if (SUCCEEDED(hr))
        {
            // load the XSLT
            VARIANT_BOOL vfSuccess = VARIANT_FALSE;
            hr = pContentXSLT->loadXML(bstrContentXSLT, &vfSuccess);
            if (!vfSuccess)
            {
                hr = FAILED(hr) ? hr : E_FAIL; // keep failed hr
            }

            if (SUCCEEDED(hr))
            {
                // get the root node of the XSLT
                IXMLDOMNode* pContentXSLTNode = NULL;
                hr = pContentXSLT->QueryInterface(IID_PPV_ARGS(&pContentXSLTNode));
                if (SUCCEEDED(hr))
                {
                    // transform the internal DOM object using the XSLT to generate the content string
                    BSTR bstrContent = NULL;
                    hr = _���Ա_XmlDom_recipe�ļ�->transformNode(pContentXSLTNode, &bstrContent);
                    if (SUCCEEDED(hr))
                    {
                        // initialize a PROPVARIANT from the string, and store it in the internal value cache
                        PROPVARIANT propvarContent = {};
                        hr = InitPropVariantFromString(bstrContent, &propvarContent);
                        if (SUCCEEDED(hr))
                        {
                            hr = _���Ա_IPropertyStoreCache_����->SetValueAndState(PKEY_Search_Contents, &propvarContent, PSC_NORMAL);
                            PropVariantClear(&propvarContent);
                        }

                        SysFreeString(bstrContent);
                    }
                    pContentXSLT->Release();
                }
            }
            pContentXSLT->Release();
        }
        SysFreeString(bstrContentXSLT);
    }

    return hr;
}

//���ڲ������е�ֵ������ڲ�DOM����
HRESULT CRecipePropertyHandler::_SaveCacheToDom()
{
    // iterate over each property in the internal value cache
    DWORD cProps;
    HRESULT hr = _���Ա_IPropertyStoreCache_����->GetCount(&cProps);
    for (UINT i = 0; SUCCEEDED(hr) && (i < cProps); ++i)
    {
        PROPERTYKEY key;
        hr = _���Ա_IPropertyStoreCache_����->GetAt(i, &key);
        if (SUCCEEDED(hr))
        {
            // check the cache state; only save dirty properties
            PSC_STATE psc;
            hr = _���Ա_IPropertyStoreCache_����->GetState(key, &psc);
            if (SUCCEEDED(hr) && (psc == PSC_DIRTY))
            {
                // get the cached value
                PROPVARIANT propvar = {};
                hr = _���Ա_IPropertyStoreCache_����->GetValue(key, &propvar);
                if (SUCCEEDED(hr))
                {
                    // save as a native property if the key is in the property map
                    BOOL fIsNativeProperty = FALSE;
                    for (UINT j = 0; j < ARRAYSIZE(ȫ��_�б�_PKEY_·��_����); ++j)
                    {
                        if (key == *ȫ��_�б�_PKEY_·��_����[j].pkey)
                        {
                            fIsNativeProperty = TRUE;
                            hr = _SaveProperty(propvar, ȫ��_�б�_PKEY_·��_����[j]);
                            break;
                        }
                    }

                    // otherwise, save as an extended proeprty
                    if (!fIsNativeProperty)
                    {
                        hr = _SaveExtendedProperty(key, propvar);
                    }

                    PropVariantClear(&propvar);
                }
            }
        }
    }

    return hr;
}

//������PROPVARIANT�е�ֵ���浽ָ����XML�ڵ�
HRESULT CRecipePropertyHandler::_SavePropertyValues(IXMLDOMNode* pNodeParent,
    PCWSTR �ڵ�����,
    REFPROPVARIANT propvarValues)
{
    // iterate through each value in the PROPVARIANT
    HRESULT hr = S_OK;
    ULONG cValues = PropVariantGetElementCount(propvarValues);
    for (ULONG iValue = 0; SUCCEEDED(hr) && (iValue < cValues); iValue++)
    {
        PROPVARIANT propvarValue = {};
        hr = PropVariantGetElem(propvarValues, iValue, &propvarValue);
        if (SUCCEEDED(hr))
        {
            // convert to a BSTR
            BSTR bstrValue = NULL;
            hr = PropVariantToBSTR(propvarValue, &bstrValue);
            if (SUCCEEDED(hr))
            {
                // create an element and set its text to the value
                IXMLDOMElement* pValue = NULL;
                hr = _���Ա_XmlDom_recipe�ļ�->createElement(CastToBSTRForInput(�ڵ�����), &pValue);
                if (SUCCEEDED(hr))
                {
                    hr = pValue->put_text(bstrValue);
                    if (SUCCEEDED(hr))
                    {
                        // append the value to its parent node
                        hr = pNodeParent->appendChild(pValue, NULL);
                    }

                    pValue->Release();
                }
                SysFreeString(bstrValue);
            }
            PropVariantClear(&propvarValue);
        }
    }
    return hr;
}

//��������PROPVARIANTֵ���浽ָ��ӳ����ָ����XML�ڵ�
HRESULT CRecipePropertyHandler::_SaveProperty(REFPROPVARIANT propvar, const PROPERTYMAP& map)
{
    // obtain the parent node of the value
    IXMLDOMNode* pNodeParent = NULL;
    HRESULT hr = _���Ա_XmlDom_recipe�ļ�->selectSingleNode(CastToBSTRForInput(map.·��), &pNodeParent);
    if (hr == S_OK)
    {
        // remove existing value nodes
        IXMLDOMNodeList* pNodeListValues = NULL;
        hr = pNodeParent->selectNodes(CastToBSTRForInput(map.�ڵ�����), &pNodeListValues);
        if (hr == S_OK)
        {
            IXMLDOMSelection* pSelectionValues = NULL;
            hr = pNodeListValues->QueryInterface(IID_PPV_ARGS(&pSelectionValues));
            if (SUCCEEDED(hr))
            {
                hr = pSelectionValues->removeAll();
                pSelectionValues->Release();
            }
            pNodeListValues->Release();
        }

        if (SUCCEEDED(hr))
        {
            // save the new values to the parent node
            hr = _SavePropertyValues(pNodeParent, map.�ڵ�����, propvar);
        }

        pNodeParent->Release();
    }
    return hr;
}

//������и�������ֵ����չ����
HRESULT CRecipePropertyHandler::_SaveExtendedProperty(REFPROPERTYKEY key, REFPROPVARIANT propvar)
{
    // convert the key to string form; don't use canonical name, since it may not be registered on other systems
    WCHAR szKey[MAX_PATH] = {};
    HRESULT hr = PSStringFromPropertyKey(key, szKey, ARRAYSIZE(szKey));
    if (SUCCEEDED(hr))
    {
        // serialize the value to string form
        PWSTR pszValue;
        hr = SerializePropVariantAsString(propvar, &pszValue);
        if (SUCCEEDED(hr))
        {
            // obtain or create the ExtendedProperties node in the document
            IXMLDOMElement* pRecipe = NULL;
            hr = _���Ա_XmlDom_recipe�ļ�->get_documentElement(&pRecipe);
            if (hr == S_OK)
            {
                IXMLDOMNode* pExtended = NULL;
                hr = _EnsureChildNodeExists(pRecipe, L"ExtendedProperties", NULL, &pExtended);
                if (SUCCEEDED(hr))
                {
                    // query for the Property node with the given key, or create a new Property node
                    WCHAR szPropertyPath[MAX_PATH];
                    hr = StringCchPrintfW(szPropertyPath, ARRAYSIZE(szPropertyPath), L"Property[@Key = '%s']", szKey);
                    if (SUCCEEDED(hr))
                    {
                        if (propvar.vt != VT_EMPTY)
                        {
                            IXMLDOMNode* pPropNode = NULL;
                            hr = _EnsureChildNodeExists(pExtended, L"Property", szPropertyPath, &pPropNode);
                            if (SUCCEEDED(hr))
                            {
                                IXMLDOMElement* pPropNodeElem = NULL;
                                hr = pPropNode->QueryInterface(IID_PPV_ARGS(&pPropNodeElem));
                                if (SUCCEEDED(hr))
                                {
                                    // set the attributes of the node with the name and value of the property
                                    VARIANT varKey = {};
                                    hr = InitVariantFromString(szKey, &varKey);
                                    if (SUCCEEDED(hr))
                                    {
                                        VARIANT varValue = {};
                                        hr = InitVariantFromString(pszValue, &varValue);
                                        if (SUCCEEDED(hr))
                                        {
                                            hr = pPropNodeElem->setAttribute(CastToBSTRForInput(L"Key"), varKey);
                                            if (SUCCEEDED(hr))
                                            {
                                                hr = pPropNodeElem->setAttribute(CastToBSTRForInput(L"EncodedValue"), varValue);
                                            }
                                            VariantClear(&varValue);
                                        }
                                        VariantClear(&varKey);
                                    }
                                    pPropNodeElem->Release();
                                }
                                pPropNode->Release();
                            }
                        }
                        else
                        {
                            // VT_EMPTY means "clear the value", so remove the corresponding node
                            IXMLDOMNode* pPropNode = NULL;
                            hr = pExtended->selectSingleNode(szPropertyPath, &pPropNode);
                            if (hr == S_OK)
                            {
                                IXMLDOMNode* pRemoved;
                                hr = pExtended->removeChild(pPropNode, &pRemoved);
                                if (SUCCEEDED(hr))
                                {
                                    pRemoved->Release();
                                }
                                pPropNode->Release();
                            }
                        }
                    }
                    pExtended->Release();
                }
                pRecipe->Release();
            }
            else
            {
                hr = E_UNEXPECTED;
            }

            CoTaskMemFree(pszValue);
        }
    }

    return hr;
}

//��ѯָ�����ӽڵ㣬��������ڣ��򴴽������һ���µ��ӽڵ�
HRESULT CRecipePropertyHandler::_EnsureChildNodeExists(IXMLDOMNode* pNodeParent, PCWSTR pszName, PCWSTR pszXPath, IXMLDOMNode** ppNodeChild)
{
    // query for the child node in case it already exists
    HRESULT hr = pNodeParent->selectSingleNode(CastToBSTRForInput(pszXPath ? pszXPath : pszName), ppNodeChild);
    if (hr != S_OK)
    {
        // create an element with the specified name and append it to the given parent node
        IXMLDOMElement* pChildElem = NULL;
        hr = _���Ա_XmlDom_recipe�ļ�->createElement(CastToBSTRForInput(pszName), &pChildElem);
        if (SUCCEEDED(hr))
        {
            hr = pNodeParent->appendChild(pChildElem, ppNodeChild);
            pChildElem->Release();
        }
    }

    return hr;
}

//��һ��PROPVARIANTֵ���л�Ϊ�ַ�����ʽ
HRESULT SerializePropVariantAsString(REFPROPVARIANT propvar, PWSTR* ppszOut)
{
    SERIALIZEDPROPERTYVALUE* pBlob;
    ULONG cbBlob;

    // serialize PROPVARIANT to binary form
    HRESULT hr = StgSerializePropVariant(&propvar, &pBlob, &cbBlob);
    if (SUCCEEDED(hr))
    {
        // determine the required buffer size
        DWORD cchString;
        hr = CryptBinaryToStringW((BYTE*)pBlob, cbBlob, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, NULL, &cchString) ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            // allocate a sufficient buffer
            PWSTR pszOut = (PWSTR)CoTaskMemAlloc(sizeof(WCHAR) * cchString);
            hr = pszOut ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                // convert the serialized binary blob to a string representation
                hr = CryptBinaryToStringW((BYTE*)pBlob, cbBlob, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, pszOut, &cchString) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    *ppszOut = pszOut;
                    pszOut = NULL; // ownership transferred to *ppszOut
                }
                CoTaskMemFree(pszOut);
            }
        }
        CoTaskMemFree(pBlob);
    }
    return hr;
}

//���ַ���ֵ�����л�ΪPROPVARIANT��ʽ
HRESULT DeserializePropVariantFromString(PCWSTR pszIn, PROPVARIANT* ppropvar)
{
    HRESULT hr = E_FAIL;
    DWORD dwFormatUsed, dwSkip, cbBlob;

    // compute and validate the required buffer size
    if (CryptStringToBinaryW(pszIn, 0, CRYPT_STRING_BASE64, NULL, &cbBlob, &dwSkip, &dwFormatUsed) &&
        dwSkip == 0 &&
        dwFormatUsed == CRYPT_STRING_BASE64)
    {
        // allocate a buffer to hold the serialized binary blob
        BYTE* pbSerialized = (BYTE*)CoTaskMemAlloc(cbBlob);
        hr = pbSerialized ? S_OK : E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            // convert the string to a serialized binary blob
            hr = CryptStringToBinaryW(pszIn, 0, CRYPT_STRING_BASE64, pbSerialized, &cbBlob, &dwSkip, &dwFormatUsed) ? S_OK : E_FAIL;
            if (SUCCEEDED(hr))
            {
                // deserialized the blob back into a PROPVARIANT value
                hr = StgDeserializePropVariant((SERIALIZEDPROPERTYVALUE*)pbSerialized, cbBlob, ppropvar);
            }
            CoTaskMemFree(pbSerialized);
        }
    }
    return hr;
}

