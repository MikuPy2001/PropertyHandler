#include <shobjidl.h>
#include <msxml6.h>

#include "Dll.h"
#include "XmlHelp.h"
#include "SafeRelease.h"

//这个类负责解析 XML

__inline BSTR CastToBSTRForInput(PCWSTR psz) { return const_cast<BSTR>(psz); }


HRESULT XmlHelp(IStream* pStream_文件流, IXMLDOMDocument** XML)
{
    OutputDebugString(L"XmlHelp()\r\n");
    HRESULT hr = E_UNEXPECTED;
    //获取XmlDom的COM对象
    hr = CoCreateInstance(
        __uuidof(DOMDocument60),
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(XML)//获取成功将由此变量回传
    );
    if (SUCCEEDED(hr))
    {
        // 从流中加载DOM对象的内容
        VARIANT varStream_xml读取 = {};
        varStream_xml读取.vt = VT_UNKNOWN;
        varStream_xml读取.punkVal = pStream_文件流;
        VARIANT_BOOL vfSuccess_xml结果 = VARIANT_FALSE;
        hr = (*XML)->load(varStream_xml读取, &vfSuccess_xml结果);//解析文件
        if (hr == S_OK && vfSuccess_xml结果 == VARIANT_TRUE)
        {
            OutputDebugString(L"XmlHelp() 文件解析成功\r\n");
        }
        else
        {
            OutputDebugString(L"XmlHelp() 文件解析失败\r\n");
            hr = E_FAIL;
        }

        if (FAILED(hr))
        {
            OutputDebugString(L"XmlHelp() XML 被 SafeRelease\r\n");
            SafeRelease(XML);
        }
    }
    return hr;
}

HRESULT XmlHelp_getV(IXMLDOMDocument* XML, PCWSTR 父路径, PCWSTR 节点名, BSTR** 子节点文本数组, long* 子节点数量)
{
    OutputDebugString(L"XmlHelp::getV() ");
    OutputDebugString(父路径);
    OutputDebugString(L" ");
    OutputDebugString(节点名);
    OutputDebugString(L"\r\n");
    // 选择属性的父节点并加载属性的值
    IXMLDOMNode* 父节点 = NULL;
    HRESULT hr = XML->selectSingleNode(CastToBSTRForInput(父路径), &父节点);
    if (hr == S_OK)
    {
        PROPVARIANT propvarValues_xml值 = {};

        // select the value nodes
        IXMLDOMNodeList* 子节点列表 = NULL;
        hr = 父节点->selectNodes(CastToBSTRForInput(节点名), &子节点列表);

        if (hr == S_OK)
        {
            // 获取值的计数
            *子节点数量 = 0;
            hr = 子节点列表->get_length(子节点数量);
            if (SUCCEEDED(hr))
            {
                // 创建一个数组来保存这些值
                // wchar_t**
                BSTR* 子节点文本数组_t = new(std::nothrow) BSTR[*子节点数量];
                hr = 子节点文本数组_t ? S_OK : E_OUTOFMEMORY;
                *子节点文本数组 = 子节点文本数组_t;
                if (SUCCEEDED(hr))
                {
                    // 将每个值节点的文本加载到数组中
                    for (long 计次 = 0; 计次 < *子节点数量; 计次++)
                    {
                        if (hr == S_OK)
                        {
                            IXMLDOMNode* 子节点 = NULL;
                            hr = 子节点列表->get_item(计次, &子节点);
                            if (hr == S_OK)
                            {
                                hr = 子节点->get_text(&子节点文本数组_t[计次]);
                                OutputDebugString(子节点文本数组_t[计次]);
                                子节点->Release();
                            }
                        }
                        else
                        {
                            //失败时将剩余的元素置空
                            子节点文本数组_t[计次] = NULL;
                        }
                    }
                }
            }
        }
        父节点->Release();
    }
    return hr;
}

HRESULT XmlHelp_getV_Release(BSTR* 子节点文本数组, long 子节点数量)
{
    // 清理数组的值
    for (long iValue = 0; iValue < 子节点数量; iValue++)
    {
        SysFreeString(子节点文本数组[iValue]);
    }
    delete[] 子节点文本数组;

    return S_OK;
}