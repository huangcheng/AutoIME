//
// Created by cheng on 2024-04-20.
//

#ifndef AUTOIME_LIBRARY_H
#define AUTOIME_LIBRARY_H
#include <windows.h>

extern "C" {
   bool GetIMEs(_Inout_ LPTSTR lpBuffer, _In_ size_t numberOfElements);

   bool SetIME(_In_ LPCTSTR name);
};

#endif //AUTOIME_LIBRARY_H
