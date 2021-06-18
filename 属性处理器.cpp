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

// 等价于 ppt.Release();
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
    PCWSTR 路径;
    PCWSTR 节点名称;
};

const PROPERTYMAP 全局_列表_PKEY_路径_名称[] =
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
    CRecipePropertyHandler() : _cRef(1), _类成员_文件流(NULL), _grfMode_给Initialize的STGM模式(0), _类成员_XmlDom_recipe文件(NULL), _类成员_IPropertyStoreCache_缓存(NULL)
    {
        DllAddRef();//每次初始化都加一次锁
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
    // 此接口公开用于枚举和操作属性值的方法。
    IFACEMETHODIMP GetCount(DWORD* pcProps);//此方法返回附加到文件的属性数的计数。
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey);//从项的属性数组中获取属性键。
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

    //每个class被销毁都
    ~CRecipePropertyHandler()
    {
        SafeRelease(&_类成员_文件流);
        SafeRelease(&_类成员_XmlDom_recipe文件);
        SafeRelease(&_类成员_IPropertyStoreCache_缓存);
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
    IStream* _类成员_文件流; // 传入Initialize的数据流，并在提交时保存
    DWORD                _grfMode_给Initialize的STGM模式; // 传递给Initialize的STGM模式
    IXMLDOMDocument* _类成员_XmlDom_recipe文件; // 表示.recipe文件的DOM对象
    IPropertyStoreCache* _类成员_IPropertyStoreCache_缓存;  // 内部值缓存从DOM后端抽象IPropertyStore操作
};
// 负责实例化对象并推送给系统,然后销毁
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
//访问方法直接转发到内部值缓存
HRESULT CRecipePropertyHandler::GetCount(DWORD* pcProps){    *pcProps = 0;    return _类成员_IPropertyStoreCache_缓存 ? _类成员_IPropertyStoreCache_缓存->GetCount(pcProps) : E_UNEXPECTED;}

HRESULT CRecipePropertyHandler::GetAt(DWORD iProp, PROPERTYKEY* pkey)
{
    *pkey = PKEY_Null;
    return _类成员_IPropertyStoreCache_缓存 ? _类成员_IPropertyStoreCache_缓存->GetAt(iProp, pkey) : E_UNEXPECTED;
}

HRESULT CRecipePropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar)
{
    PropVariantInit(pPropVar);
    return _类成员_IPropertyStoreCache_缓存 ? _类成员_IPropertyStoreCache_缓存->GetValue(key, pPropVar) : E_UNEXPECTED;
}

// SetValue只是更新内部值缓存
HRESULT CRecipePropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
    HRESULT hr = E_UNEXPECTED;
    if (_类成员_IPropertyStoreCache_缓存)
    {
        // check grfMode to ensure writes are allowed
        hr = STG_E_ACCESSDENIED;
        if ((_grfMode_给Initialize的STGM模式 & STGM_READWRITE) &&
            (key != PKEY_Search_Contents))  // this property is read-only
        {
            hr = _类成员_IPropertyStoreCache_缓存->SetValueAndState(key, &propVar, PSC_DIRTY);
        }
    }

    return hr;
}

// Commit将内部值缓存写回传递给Initialize的流
HRESULT CRecipePropertyHandler::Commit()
{
    HRESULT hr = E_UNEXPECTED;
    if (_类成员_IPropertyStoreCache_缓存)
    {
        // 检查grfMode，确保允许写
        hr = STG_E_ACCESSDENIED;
        if (_grfMode_给Initialize的STGM模式 & STGM_READWRITE)
        {
            // 将内部值缓存保存到XML DOM对象
            hr = _SaveCacheToDom();
            if (SUCCEEDED(hr))
            {
                // 重置输出流
                LARGE_INTEGER liZero = {};
                hr = _类成员_文件流->Seek(liZero, STREAM_SEEK_SET, NULL);
                if (SUCCEEDED(hr))
                {
                    // 获取临时目标流进行手动安全保存
                    IDestinationStreamFactory* pSafeCommit;
                    hr = _类成员_文件流->QueryInterface(&pSafeCommit);
                    if (SUCCEEDED(hr))
                    {
                        IStream* pStreamCommit;
                        hr = pSafeCommit->GetDestinationStream(&pStreamCommit);
                        //获取一个空流，该流接收要复制的文件的新版本。
                        if (SUCCEEDED(hr))
                        {
                            // 将XML写入临时流并提交它
                            VARIANT varStream = {};
                            varStream.vt = VT_UNKNOWN;
                            varStream.punkVal = pStreamCommit;
                            hr = _类成员_XmlDom_recipe文件->save(varStream);
                            if (SUCCEEDED(hr))
                            {
                                hr = pStreamCommit->Commit(STGC_DEFAULT);
                                if (SUCCEEDED(hr))
                                {
                                    // 提交真实的输出流
                                    _类成员_文件流->Commit(STGC_DEFAULT);
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

//指示用户是否应该能够编辑给定属性键的值
HRESULT CRecipePropertyHandler::IsPropertyWritable(REFPROPERTYKEY key)
{
    // System.Search.Contents 是唯一不支持写入的属性
    return (key == PKEY_Search_Contents) ? S_FALSE : S_OK;
}

// Initialize使用来自指定流的数据填充内部值缓存
HRESULT CRecipePropertyHandler::Initialize(IStream* pStream_文件流, DWORD grfMode)
{
    HRESULT hr = E_UNEXPECTED;
    if (!_类成员_文件流)//禁止初始化两次?
    {
        //获取XmlDom的COM对象
        hr = CoCreateInstance(
            __uuidof(DOMDocument60), 
            NULL, 
            CLSCTX_INPROC_SERVER, 
            IID_PPV_ARGS(&_类成员_XmlDom_recipe文件)//获取成功将由此变量回传
        );
        if (SUCCEEDED(hr))
        {
            // 从流中加载DOM对象的内容
            VARIANT varStream_xml读取 = {};
            varStream_xml读取.vt = VT_UNKNOWN;
            varStream_xml读取.punkVal = pStream_文件流;
            VARIANT_BOOL vfSuccess_xml结果 = VARIANT_FALSE;
            hr = _类成员_XmlDom_recipe文件->load(varStream_xml读取, &vfSuccess_xml结果);//解析文件
            if (hr == S_OK && vfSuccess_xml结果 == VARIANT_TRUE)
            {
                // 从DOM对象加载内部值缓存
                hr = _LoadCacheFromDom();
                if (SUCCEEDED(hr))
                {
                    // 保存对流和grfMode的引用
                    hr = pStream_文件流->QueryInterface(IID_PPV_ARGS(&_类成员_文件流));
                    if (SUCCEEDED(hr))
                    {
                        _grfMode_给Initialize的STGM模式 = grfMode;
                    }
                }
            }
            else
            {
                hr = E_FAIL;
            }

            if (FAILED(hr))
            {
                SafeRelease(&_类成员_XmlDom_recipe文件);
            }
        }
    }

    return hr;
}

//从内部DOM对象填充内部值缓存
//解析 _类成员_XmlDom_recipe文件
//存放到 cache
HRESULT CRecipePropertyHandler::_LoadCacheFromDom()
{
    HRESULT hr = S_OK;

    if (!_类成员_IPropertyStoreCache_缓存)
    {
        // 创建内部值缓存
        hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&_类成员_IPropertyStoreCache_缓存));//初始化 内存属性存储
        if (SUCCEEDED(hr))
        {
            // 直接从XML填充本机属性
            for (UINT i = 0; i < ARRAYSIZE(全局_列表_PKEY_路径_名称); ++i)
            {
                _LoadProperty(全局_列表_PKEY_路径_名称[i]);
            }

            // 加载扩展属性和搜索内容
            _LoadExtendedProperties();
            _LoadSearchContent();
        }
    }

    return hr;
}

//从给定的父节点加载指定的值，并从它们创建PROPVARIANT
HRESULT CRecipePropertyHandler::_LoadPropertyValues(IXMLDOMNode* 父节点,
    PCWSTR 子节点名称,
    PROPVARIANT* 返回值)
{
    // 初始化
    PropVariantInit(返回值);

    // select the value nodes
    IXMLDOMNodeList* 子节点列表 = NULL;
    HRESULT hr = 父节点->selectNodes(CastToBSTRForInput(子节点名称), &子节点列表);
    if (hr == S_OK)
    {
        // 获取值的计数
        long 子节点数量 = 0;
        hr = 子节点列表->get_length(&子节点数量);
        if (SUCCEEDED(hr))
        {
            // 创建一个数组来保存这些值
            // wchar_t**
            BSTR* 子节点文本数组 = new(std::nothrow) BSTR[子节点数量];
            hr = 子节点文本数组 ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                // 将每个值节点的文本加载到数组中
                for (long 计次 = 0; 计次 < 子节点数量; 计次++)
                {
                    if (hr == S_OK)
                    {
                        IXMLDOMNode* 子节点 = NULL;
                        hr = 子节点列表->get_item(计次, &子节点);
                        if (hr == S_OK)
                        {
                            hr = 子节点->get_text(&子节点文本数组[计次]);
                            子节点->Release();
                        }
                    }
                    else
                    {
                        //失败时将剩余的元素置空
                        子节点文本数组[计次] = NULL;
                    }
                }

                if (hr == S_OK)
                {
                    // 将值的列表打包到PROPVARIANT中
                    hr = InitPropVariantFromStringVector(const_cast<PCWSTR*>(子节点文本数组), 子节点数量, 返回值);
                }

                // 清理数组的值
                for (long iValue = 0; iValue < 子节点数量; iValue++)
                {
                    SysFreeString(子节点文本数组[iValue]);
                }
                delete[] 子节点文本数组;
            }
        }
    }

    return hr;
}

//加载给定映射中指定的属性的数据到内部值缓存中
HRESULT CRecipePropertyHandler::_LoadProperty(const PROPERTYMAP& PKEY_路径_名称)
{
    // 选择属性的父节点并加载属性的值
    IXMLDOMNode* 父节点 = NULL;
    HRESULT hr = _类成员_XmlDom_recipe文件->selectSingleNode(CastToBSTRForInput(PKEY_路径_名称.路径), &父节点);
    if (hr == S_OK)
    {
        PROPVARIANT propvarValues_xml值 = {};
        hr = _LoadPropertyValues(父节点, PKEY_路径_名称.节点名称, &propvarValues_xml值);
        if (hr == S_OK)
        {
            // 强制将值转换为属性键的适当类型
            hr = PSCoerceToCanonicalValue(*PKEY_路径_名称.pkey, &propvarValues_xml值);
            if (SUCCEEDED(hr))
            {
                // 缓存加载的值
                hr = _类成员_IPropertyStoreCache_缓存->SetValueAndState(*PKEY_路径_名称.pkey, &propvarValues_xml值, PSC_NORMAL);
                // PSC_NORMAL 财产未被改变。
                // PSC_NOTINSOURCE	初始化属性处理程序的文件或流的请求属性不存在。
                // PSC_DIRTY	属性已被更改，但尚未提交到文件或流。
                // PSC_READONLY
            }
            PropVariantClear(&propvarValues_xml值);
        }
        父节点->Release();
    }

    return hr;
}

//加载任何外部属性的数据(例如那些没有显式映射到XML模式的属性)到内部值缓存
HRESULT CRecipePropertyHandler::_LoadExtendedProperties()
{
    // select the list of extended property nodes
    IXMLDOMNodeList* pList = NULL;
    HRESULT hr = _类成员_XmlDom_recipe文件->selectNodes(CastToBSTRForInput(L"Recipe/ExtendedProperties/Property"), &pList);
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
                                        hr = _类成员_IPropertyStoreCache_缓存->SetValueAndState(key, &propvarValue, PSC_NORMAL);
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

//填充System.Search。内部值缓存中的内容属性
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
                    hr = _类成员_XmlDom_recipe文件->transformNode(pContentXSLTNode, &bstrContent);
                    if (SUCCEEDED(hr))
                    {
                        // initialize a PROPVARIANT from the string, and store it in the internal value cache
                        PROPVARIANT propvarContent = {};
                        hr = InitPropVariantFromString(bstrContent, &propvarContent);
                        if (SUCCEEDED(hr))
                        {
                            hr = _类成员_IPropertyStoreCache_缓存->SetValueAndState(PKEY_Search_Contents, &propvarContent, PSC_NORMAL);
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

//将内部缓存中的值保存回内部DOM对象
HRESULT CRecipePropertyHandler::_SaveCacheToDom()
{
    // iterate over each property in the internal value cache
    DWORD cProps;
    HRESULT hr = _类成员_IPropertyStoreCache_缓存->GetCount(&cProps);
    for (UINT i = 0; SUCCEEDED(hr) && (i < cProps); ++i)
    {
        PROPERTYKEY key;
        hr = _类成员_IPropertyStoreCache_缓存->GetAt(i, &key);
        if (SUCCEEDED(hr))
        {
            // check the cache state; only save dirty properties
            PSC_STATE psc;
            hr = _类成员_IPropertyStoreCache_缓存->GetState(key, &psc);
            if (SUCCEEDED(hr) && (psc == PSC_DIRTY))
            {
                // get the cached value
                PROPVARIANT propvar = {};
                hr = _类成员_IPropertyStoreCache_缓存->GetValue(key, &propvar);
                if (SUCCEEDED(hr))
                {
                    // save as a native property if the key is in the property map
                    BOOL fIsNativeProperty = FALSE;
                    for (UINT j = 0; j < ARRAYSIZE(全局_列表_PKEY_路径_名称); ++j)
                    {
                        if (key == *全局_列表_PKEY_路径_名称[j].pkey)
                        {
                            fIsNativeProperty = TRUE;
                            hr = _SaveProperty(propvar, 全局_列表_PKEY_路径_名称[j]);
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

//将给定PROPVARIANT中的值保存到指定的XML节点
HRESULT CRecipePropertyHandler::_SavePropertyValues(IXMLDOMNode* pNodeParent,
    PCWSTR 节点名称,
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
                hr = _类成员_XmlDom_recipe文件->createElement(CastToBSTRForInput(节点名称), &pValue);
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

//将给定的PROPVARIANT值保存到指定映射所指定的XML节点
HRESULT CRecipePropertyHandler::_SaveProperty(REFPROPVARIANT propvar, const PROPERTYMAP& map)
{
    // obtain the parent node of the value
    IXMLDOMNode* pNodeParent = NULL;
    HRESULT hr = _类成员_XmlDom_recipe文件->selectSingleNode(CastToBSTRForInput(map.路径), &pNodeParent);
    if (hr == S_OK)
    {
        // remove existing value nodes
        IXMLDOMNodeList* pNodeListValues = NULL;
        hr = pNodeParent->selectNodes(CastToBSTRForInput(map.节点名称), &pNodeListValues);
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
            hr = _SavePropertyValues(pNodeParent, map.节点名称, propvar);
        }

        pNodeParent->Release();
    }
    return hr;
}

//保存具有给定键和值的扩展属性
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
            hr = _类成员_XmlDom_recipe文件->get_documentElement(&pRecipe);
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

//查询指定的子节点，如果不存在，则创建并添加一个新的子节点
HRESULT CRecipePropertyHandler::_EnsureChildNodeExists(IXMLDOMNode* pNodeParent, PCWSTR pszName, PCWSTR pszXPath, IXMLDOMNode** ppNodeChild)
{
    // query for the child node in case it already exists
    HRESULT hr = pNodeParent->selectSingleNode(CastToBSTRForInput(pszXPath ? pszXPath : pszName), ppNodeChild);
    if (hr != S_OK)
    {
        // create an element with the specified name and append it to the given parent node
        IXMLDOMElement* pChildElem = NULL;
        hr = _类成员_XmlDom_recipe文件->createElement(CastToBSTRForInput(pszName), &pChildElem);
        if (SUCCEEDED(hr))
        {
            hr = pNodeParent->appendChild(pChildElem, ppNodeChild);
            pChildElem->Release();
        }
    }

    return hr;
}

//将一个PROPVARIANT值序列化为字符串形式
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

//将字符串值反序列化为PROPVARIANT形式
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

