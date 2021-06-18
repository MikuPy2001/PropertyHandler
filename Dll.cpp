// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

#include <windows.h>
#include <new>  // std::nothrow
#include <shlobj.h>
#include <shlwapi.h>
#include "Dll.h"
#include  <Locale.h>


// Handle the the DLL's module
HINSTANCE g_hInst = NULL;

// Standard DLL functions
STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, void*)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
        setlocale(LC_CTYPE, "");
        g_hInst = hInstance;
        DisableThreadLibraryCalls(hInstance);
    }
    return TRUE;
}

long g_cRefModule = 0;
STDAPI DllCanUnloadNow()
{
    // ֻ�������ͷ�����δ��ɵ����ú�ж��DLL
    return (g_cRefModule == 0) ? S_OK : S_FALSE;
}

//����
void DllAddRef()
{
    InterlockedIncrement(&g_cRefModule);
}
//����
void DllRelease()
{
    InterlockedDecrement(&g_cRefModule);
}


