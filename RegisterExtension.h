// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

#pragma once

#include <windows.h>

// 生产代码应该使用安装程序技术，如MSI注册它的处理程序，而不是这个类。
// 这个类用于演示，它封装了不同类型的处理程序注册，通过提供具有映射到所支持的扩展模式的参数的方法来对这些注册进行schemalize，并使创建自注册的.exe和.dll变得容易

class CRegisterExtension
{
public:
    CRegisterExtension(REFCLSID clsid = CLSID_NULL, HKEY hkeyRoot = HKEY_CURRENT_USER);
    ~CRegisterExtension();
    void SetHandlerCLSID(REFCLSID clsid);
    void SetInstallScope(HKEY hkeyRoot);
    HRESULT SetModule(PCWSTR pszModule);
    HRESULT SetModule(HINSTANCE hinst);

    HRESULT RegisterInProcServer(PCWSTR pszFriendlyName, PCWSTR pszThreadingModel) const;
    HRESULT RegisterInProcServerAttribute(PCWSTR pszAttribute, DWORD dwValue) const;

    // 为当前运行的模块注册COM本地服务器这是用于自注册应用程序的
    // 注册应用程序为本地服务器
    HRESULT RegisterAppAsLocalServer(PCWSTR pszFriendlyName, PCWSTR pszCmdLine = NULL) const;
    // 注册可提升的本地服务器
    HRESULT RegisterElevatableLocalServer(PCWSTR pszFriendlyName, UINT idLocalizeString, UINT idIconRef) const;
    // 在Proc服务器中注册Elevatable
    HRESULT RegisterElevatableInProcServer(PCWSTR pszFriendlyName, UINT idLocalizeString, UINT idIconRef) const;

    // 移除COM对象注册
    HRESULT UnRegisterObject() const;

    // 这允许直接拖放到。exe上，如果你有exe的快捷方式(或者可以通过send to菜单访问。exe)，这很有用。

    HRESULT RegisterAppDropTarget() const;

    // 为基于静态谓词的删除目标创建注册表项。
    // 指定的clsid将是

    HRESULT RegisterCreateProcessVerb(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszCmdLine, PCWSTR pszVerbDisplayName) const;
    HRESULT RegisterDropTargetVerb(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszVerbDisplayName) const;
    HRESULT RegisterExecuteCommandVerb(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszVerbDisplayName) const;
    HRESULT RegisterExplorerCommandVerb(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszVerbDisplayName) const;
    HRESULT RegisterExplorerCommandStateHandler(PCWSTR pszProgID, PCWSTR pszVerb) const;
    HRESULT RegisterVerbAttribute(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszValueName) const;
    HRESULT RegisterVerbAttribute(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszValueName, PCWSTR pszValue) const;
    HRESULT RegisterVerbAttribute(PCWSTR pszProgID, PCWSTR pszVerb, PCWSTR pszValueName, DWORD dwValue) const;
    HRESULT RegisterVerbDefaultAndOrder(PCWSTR pszProgID, PCWSTR pszVerbOrderFirstIsDefault) const;

    HRESULT RegisterPlayerVerbs(PCWSTR const rgpszAssociation[], UINT countAssociation,
                                PCWSTR pszVerb, PCWSTR pszTitle) const;

    HRESULT UnRegisterVerb(PCWSTR pszProgID, PCWSTR pszVerb) const;
    HRESULT UnRegisterVerbs(PCWSTR const rgpszAssociation[], UINT countAssociation, PCWSTR pszVerb) const;

    HRESULT RegisterContextMenuHandler(PCWSTR pszProgID, PCWSTR pszDescription) const;
    HRESULT RegisterRightDragContextMenuHandler(PCWSTR pszProgID, PCWSTR pszDescription) const;

    HRESULT RegisterAppShortcutInSendTo() const;

    HRESULT RegisterThumbnailHandler(PCWSTR pszExtension) const;
    HRESULT RegisterPropertyHandler(PCWSTR pszExtension) const;
    HRESULT UnRegisterPropertyHandler(PCWSTR pszExtension) const;

    HRESULT RegisterLinkHandler(PCWSTR pszProgID) const;

    HRESULT RegisterProgID(PCWSTR pszProgID, PCWSTR pszTypeName, UINT idIcon) const;
    HRESULT UnRegisterProgID(PCWSTR pszProgID, PCWSTR pszFileExtension) const;
    HRESULT RegisterExtensionWithProgID(PCWSTR pszFileExtension, PCWSTR pszProgID) const;
    HRESULT RegisterOpenWith(PCWSTR pszFileExtension, PCWSTR pszProgID) const;
    HRESULT RegisterNewMenuNullFile(PCWSTR pszFileExtension, PCWSTR pszProgID) const;
    HRESULT RegisterNewMenuData(PCWSTR pszFileExtension, PCWSTR pszProgID, PCSTR pszBase64) const;
    HRESULT RegisterKind(PCWSTR pszFileExtension, PCWSTR pszKindValue) const;
    HRESULT UnRegisterKind(PCWSTR pszFileExtension) const;
    HRESULT RegisterPropertyHandlerOverride(PCWSTR pszProperty) const;

    HRESULT RegisterHandlerSupportedProtocols(PCWSTR pszProtocol) const;

    HRESULT RegisterProgIDValue(PCWSTR pszProgID, PCWSTR pszValueName) const;
    HRESULT RegisterProgIDValue(PCWSTR pszProgID, PCWSTR pszValueName, PCWSTR pszValue) const;
    HRESULT RegisterProgIDValue(PCWSTR pszProgID, PCWSTR pszValueName, DWORD dwValue) const;

    // 这可能应该是私有的，但它们是有用的
    HRESULT RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, PCWSTR pszValue, ...) const;
    HRESULT RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, DWORD dwValue, ...) const;
    HRESULT RegSetKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, const unsigned char pc[], DWORD dwSize, ...) const;
    HRESULT RegSetKeyValueBinaryPrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValueName, PCSTR pszBase64, ...) const;

    HRESULT RegDeleteKeyPrintf(HKEY hkey, PCWSTR pszKeyFormatString, ...) const;
    HRESULT RegDeleteKeyValuePrintf(HKEY hkey, PCWSTR pszKeyFormatString, PCWSTR pszValue, ...) const;

    PCWSTR GetCLSIDString() const { return _szCLSID; };

    bool HasClassID() const { return _clsid != CLSID_NULL ? true : false; };

private:

    HRESULT _EnsureModule() const;
    bool _IsBaseClassProgID(PCWSTR pszProgID)  const;
    HRESULT _EnsureBaseProgIDVerbIsNone(PCWSTR pszProgID) const;
    void _UpdateAssocChanged(HRESULT hr, PCWSTR pszKeyFormatString) const;

    CLSID _clsid;
    HKEY _hkeyRoot;
    WCHAR _szCLSID[39];
    WCHAR _szModule[MAX_PATH];
    bool _fAssocChanged;
};
