#pragma once

// µÈ¼ÛÓÚ ppt.Release();
template <class T> void SafeRelease(T** ppT)
{
    if (*ppT)
    {
        OutputDebugString(L"SafeRelease()\r\n");
        (*ppT)->Release();
        *ppT = NULL;
    }
}