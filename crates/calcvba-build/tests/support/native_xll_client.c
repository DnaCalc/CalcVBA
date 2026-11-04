#include <windows.h>
#include <stdio.h>
#include <stddef.h>
typedef struct { double number; LONGLONG unused1, unused2; LONG type, padding; } Oper;
typedef double (*Binary)(double,double);
typedef double (*Unary)(double);
typedef Oper* (*Twice)(Oper*);
typedef void (*Free)(Oper*);
typedef LONG (*Count)(void);
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
