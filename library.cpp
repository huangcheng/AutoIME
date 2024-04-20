#include <msctf.h>
#include <tchar.h>

#include "library.h"

bool GetIMEs(_Inout_ LPTSTR lpBuffer, _In_ size_t numberOfElements)
{
    bool result = false;
    bool initialized = false;

    ITfInputProcessorProfiles* lpProfiles = nullptr;
    ITfInputProcessorProfileMgr* lpMgr = nullptr;
    IEnumTfInputProcessorProfiles* lpEnum = nullptr;

    __try
    {
        HRESULT hr = CoInitialize(nullptr);

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = CoCreateInstance(
                CLSID_TF_InputProcessorProfiles,
                nullptr,
                CLSCTX_INPROC_SERVER,
                IID_ITfInputProcessorProfileMgr,
                (LPVOID*)&lpMgr
        );

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = CoCreateInstance(
                CLSID_TF_InputProcessorProfiles,
                nullptr,
                CLSCTX_INPROC_SERVER,
                IID_ITfInputProcessorProfiles,
                (LPVOID*)&lpProfiles
        );

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = lpMgr->EnumProfiles(0, &lpEnum);

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        TF_INPUTPROCESSORPROFILE profile = { 0 };

        ULONG fetched = 0;

        while (lpEnum->Next(1, &profile, &fetched) == S_OK)
        {
            BSTR bstrDest = nullptr;
            BOOL enabled = false;

            hr = lpProfiles->IsEnabledLanguageProfile(
                    profile.clsid,
                    profile.langid,
                    profile.guidProfile,
                    &enabled
            );

            if (FAILED(hr))
            {
                result = false;

                __leave;
            }

            if (!enabled)
            {
                continue;
            }

            hr = lpProfiles->GetLanguageProfileDescription(
                    profile.clsid,
                    profile.langid,
                    profile.guidProfile,
                    &bstrDest
            );

            if (SUCCEEDED(hr))
            {
                if (_tcslen(lpBuffer) + _tcslen(bstrDest) + 1 < numberOfElements)
                {
                    _tcscat_s(lpBuffer, numberOfElements, bstrDest);
                    _tcscat_s(lpBuffer, numberOfElements, TEXT("|"));
                }

                SysFreeString(bstrDest);
            }

            ZeroMemory(&profile, sizeof(TF_INPUTPROCESSORPROFILE));
        }
    }
    __finally
    {
        if (lpMgr != nullptr)
        {
            lpMgr->Release();

            lpMgr = nullptr;
        }

        if (lpProfiles != nullptr)
        {
            lpProfiles->Release();

            lpProfiles = nullptr;
        }

        if (lpEnum != nullptr)
        {
            lpEnum->Release();

            lpProfiles = nullptr;
        }

        if (initialized)
        {
            CoUninitialize();
        }
    }

    if (_tcslen(lpBuffer) > 0)
    {
        auto c = _tcsrchr(lpBuffer, TEXT('|'));

        if (c != nullptr)
            {
                *c = TEXT('\0');
            }
        }

    return result;
}

bool SetIME(_In_ LPCTSTR name) {
    bool result = false;
    bool initialized = false;

    ITfInputProcessorProfiles* lpProfiles = nullptr;
    ITfInputProcessorProfileMgr* lpMgr = nullptr;
    IEnumTfInputProcessorProfiles* lpEnum = nullptr;

    __try
    {
        HRESULT hr = CoInitialize(nullptr);

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = CoCreateInstance(
            CLSID_TF_InputProcessorProfiles,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_ITfInputProcessorProfileMgr,
            (LPVOID*)&lpMgr
        );

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = CoCreateInstance(
            CLSID_TF_InputProcessorProfiles,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_ITfInputProcessorProfiles,
            (LPVOID*)&lpProfiles
        );

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = lpMgr->EnumProfiles(0, &lpEnum);

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        TF_INPUTPROCESSORPROFILE profile = { 0 };

        ULONG fetched = 0;

        while (lpEnum->Next(1, &profile, &fetched) == S_OK)
        {
            BSTR bstrDest = nullptr;

            hr = lpProfiles->GetLanguageProfileDescription(
                profile.clsid,
                profile.langid,
                profile.guidProfile,
                &bstrDest
            );

            if (SUCCEEDED(hr))
            {
                if (_tcscmp(name, bstrDest) == 0)
                {
                    hr = lpMgr->ActivateProfile(
                        TF_PROFILETYPE_INPUTPROCESSOR,
                        profile.langid,
                        profile.clsid,
                        profile.guidProfile,
                        NULL,
                        TF_IPPMF_FORSESSION | TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE
                    );

                    if (SUCCEEDED(hr))
                    {
                        result = true;

                        break;
                    }
                    else {
                        result = false;

                        __leave;
                    }

                }

                SysFreeString(bstrDest);
            }

            ZeroMemory(&profile, sizeof(TF_INPUTPROCESSORPROFILE));
        }
    }
    __finally
    {
        if (lpMgr != nullptr)
        {
            lpMgr->Release();

            lpMgr = nullptr;
        }

        if (lpProfiles != nullptr)
        {
            lpProfiles->Release();

            lpProfiles = nullptr;
        }

        if (lpEnum != nullptr)
        {
            lpEnum->Release();

            lpProfiles = nullptr;
        }

        if (initialized)
        {
            CoUninitialize();
        }
    }

    return result;
}
