#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <wchar.h>
typedef BSTR (*Suite)(void); typedef LONG (*Shutdown)(void);
int wmain(int argc,wchar_t **argv) {
 if(argc!=3) return 2; HMODULE fixture=LoadLibraryW(argv[2]),dll=LoadLibraryW(argv[1]); if(!fixture||!dll)return 3;
 Suite suite=(Suite)GetProcAddress(dll,"ForeignSuite"); Shutdown shutdown=(Shutdown)GetProcAddress(dll,"CalcVbaShutdown"); if(!suite||!shutdown)return 4;
 BSTR value=suite(); if(!value) { typedef LONG (*Error)(BSTR*,BSTR*); Error error=(Error)GetProcAddress(dll,"CalcVbaGetLastError"); BSTR source=NULL,message=NULL; LONG code=error(&source,&message); fwprintf(stderr,L"ForeignSuite: %ld %s %s\n",code,source?source:L"",message?message:L""); SysFreeString(source); SysFreeString(message); return 5; } wprintf(L"%s\n",value); SysFreeString(value); if(shutdown()!=0)return 6;
 if(!FreeLibrary(dll)||!FreeLibrary(fixture)) return 7;
 if(GetModuleHandleW(argv[2])!=NULL) return 8;
 return 0;
}
