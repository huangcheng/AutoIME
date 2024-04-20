#include <msctf.h>
#include <tchar.h>

#include "library.h"

bool GetIMEsByKeyboard(_In_ int keyborad, _Inout_ LPTSTR lpBuffer, _In_ size_t numberOfElements)
{
    bool result = true;
    bool initialized = false;

    ITfInputProcessorProfiles* lpProfiles = nullptr;

    IEnumTfLanguageProfiles* lpEnum = nullptr;

    __try
    {
        if (lpBuffer == nullptr)
        {
            result = false;

            __leave;
        }

        HRESULT hr = CoInitialize(nullptr);

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        initialized = true;

        TCHAR buf[MAX_PATH] = { 0 };

        hr = CoCreateInstance(
            CLSID_TF_InputProcessorProfiles,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_ITfInputProcessorProfiles,
            (VOID**)&lpProfiles
        );

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        hr = lpProfiles->EnumLanguageProfiles(keyborad, &lpEnum);

        if (FAILED(hr))
        {
            result = false;

            __leave;
        }

        if (lpEnum != nullptr)
        {
            TF_LANGUAGEPROFILE profile;
            ULONG fetched = 0;

            while (lpEnum->Next(1, &profile, &fetched) == S_OK)
            {
                BSTR bstrDesc = nullptr;
                BOOL bEnabled = false;

                hr = lpProfiles->IsEnabledLanguageProfile(profile.clsid, profile.langid, profile.guidProfile, &bEnabled);

                if (SUCCEEDED(hr) && bEnabled) {
                    hr = lpProfiles->GetLanguageProfileDescription(profile.clsid, profile.langid, profile.guidProfile, &bstrDesc);

                    if (SUCCEEDED(hr))
                    {
                        _tcscat_s(buf, MAX_PATH, bstrDesc);
                        _tcscat_s(buf, MAX_PATH, TEXT("|"));

                        SysFreeString(bstrDesc);
                    }
                }
            }

            auto c = _tcsrchr(buf, TEXT('|'));

            if (c != nullptr)
            {
                *c = TEXT('\0');
            }

            _tcsncpy_s(lpBuffer, numberOfElements, buf, _countof(buf));
        }
    }
    __finally
    {
        if (lpProfiles != nullptr)
        {
            lpProfiles->Release();

            lpProfiles = nullptr;
        }

        if (lpEnum != nullptr)
        {
            lpEnum->Release();

            lpEnum = nullptr;
        }

        if (initialized)
        {
            CoUninitialize();
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
