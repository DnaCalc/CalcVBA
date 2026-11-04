#include <windows.h>
#include <stdio.h>
#include <stddef.h>
typedef struct { double number; LONGLONG unused1, unused2; LONG type, padding; } Oper;
typedef double (*Binary)(double,double);
typedef double (*Unary)(double);
typedef Oper* (*Twice)(Oper*);
typedef void (*Free)(Oper*);
typedef LONG (*Count)(void);
/* Minimal host for the ordinary VBA registration/close lifecycle. */
static LONG registrations, removals;
__declspec(dllexport) int __cdecl MdCallBack12(int function,int count,Oper** args,Oper* result) {
    static WCHAR name[]={4,'t','e','s','t'};
    if(result) ZeroMemory(result,sizeof(*result));
    if(function==16393) { *(WCHAR**)result=name; result->type=2; }
    else if(function==149) { result->number=(double)++registrations; result->type=1; }
    else if(function==201) {
        if(count!=1 || args[0]->type!=1) return 2;
        ++removals; *(LONG*)result=1; result->type=4;
    } else if(function!=16384) return 2;
    return 0;
}
int main(int argc,char** argv) {
    if((argc!=2 && argc!=3) || sizeof(Oper)!=32 || offsetof(Oper,type)!=24) return 2;
    HMODULE dll=LoadLibraryA(argv[1]); if(!dll) return 3;
    Binary add=(Binary)GetProcAddress(dll,"AddDouble");
    Unary conversion=(Unary)GetProcAddress(dll,"ConversionProbe");
    Twice twice=(Twice)GetProcAddress(dll,"NativeTwice");
    Free release=(Free)GetProcAddress(dll,"xlAutoFree12");
    Count count=(Count)GetProcAddress(dll,"OutstandingResults");
    if(!add||!conversion||!twice||!release||!count) return 4;
    Oper input={4.5,0,0,1,0};
    printf("%.2f|%.2f|",add(2.25,3.5),conversion(3.75));
    Oper* result=twice(&input); if(!result) return 5;
    printf("%.0f|%ld|%ld|",result->number,result->type,count());
    release(result); printf("%ld\n",count());
    input.type=2; result=twice(&input); if(!result) return 6;
    if(result->type!=(16|16384) || *(LONGLONG*)result!=15) return 7;
    release(result); if(count()!=0) return 8;
    Count open=(Count)GetProcAddress(dll,"xlAutoOpen");
    Count close=(Count)GetProcAddress(dll,"xlAutoClose");
    if(!open || !close || open()!=1 || close()!=1 || registrations!=7 || removals!=7) return 11;
    if(close()!=1 || removals!=7 || open()!=1 || close()!=1 || removals!=14) return 12;
    if(argc==3) {
        LARGE_INTEGER frequency,begin,end; double value=0;
        QueryPerformanceFrequency(&frequency); QueryPerformanceCounter(&begin);
        for(int i=0;i<100000;i++) value=add(2.25,3.5);
        QueryPerformanceCounter(&end);
        printf("AddDouble_us=%.3f value=%.2f\n",(double)(end.QuadPart-begin.QuadPart)*1e6/(double)frequency.QuadPart/100000,value);
        QueryPerformanceCounter(&begin);
        for(int i=0;i<10000;i++) value=conversion(3.75);
        QueryPerformanceCounter(&end);
        printf("ConversionProbe_us=%.3f value=%.2f\n",(double)(end.QuadPart-begin.QuadPart)*1e6/(double)frequency.QuadPart/10000,value);
        input.type=1; QueryPerformanceCounter(&begin);
        for(int i=0;i<10000;i++) { result=twice(&input); if(!result) return 9; release(result); }
        QueryPerformanceCounter(&end);
        printf("NativeTwice_and_free_us=%.3f outstanding=%ld\n",(double)(end.QuadPart-begin.QuadPart)*1e6/(double)frequency.QuadPart/10000,count());
    }
    Count shutdown=(Count)GetProcAddress(dll,"CalcVbaShutdown");
    if(!shutdown || shutdown()!=0) return 10;
    FreeLibrary(dll); return 0;
}
