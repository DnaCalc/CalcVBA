#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <string.h>
#include <stddef.h>
typedef struct Oper Oper;
typedef struct Rect { LONG r0,r1,c0,c1; } Rect;
struct Oper {
    union {
        double number;
        WCHAR *text;
        LONG integer;
        struct { Oper *values; LONG rows, columns; } multi;
        struct { void *areas; LONGLONG sheet; } ref;
        struct { WORD count, padding; Rect area; } sref;
        LONGLONG words[3];
    } value;
    LONG type, padding;
};
typedef Oper *(*Echo)(const Oper*);
typedef void (*Free)(Oper*);
typedef LONG (*Count)(void);
typedef LONG (*Error)(void*,void*);
#define CHECK(t) do {if(!(t)){fprintf(stderr,"line %d: %s\n",__LINE__,#t);return 2;}}while(0)
int main(int argc,char **argv) {
    CHECK((argc==2 || argc==3) && sizeof(Oper)==32 && offsetof(Oper,type)==24);
    HMODULE dll=LoadLibraryA(argv[1]); CHECK(dll);
    Echo echo=(Echo)GetProcAddress(dll,"NativeEcho");
    Free release=(Free)GetProcAddress(dll,"xlAutoFree12");
    Count owned=(Count)GetProcAddress(dll,"OwnedBuffers");
    Count count=(Count)GetProcAddress(dll,"OutstandingResults");
    Error error=(Error)GetProcAddress(dll,"CalcVbaGetLastError");
    CHECK(echo && release && owned && count);
    typedef Oper *(*VariantEcho)(VARIANT*);
    VariantEcho variant_echo=(VariantEcho)GetProcAddress(dll,"NativeVariantEcho");CHECK(variant_echo);
    VARIANT variants[9]={0};double expected[9]={7,-123,456,9007199254740992.0,1.25,2.75,12.3456,45123.25,1.125};
    V_VT(&variants[0])=VT_UI1;V_UI1(&variants[0])=7;
    V_VT(&variants[1])=VT_I2;V_I2(&variants[1])=-123;
    V_VT(&variants[2])=VT_I4;V_I4(&variants[2])=456;
    V_VT(&variants[3])=VT_I8;V_I8(&variants[3])=9007199254740993LL;
    V_VT(&variants[4])=VT_R4;V_R4(&variants[4])=1.25f;
    V_VT(&variants[5])=VT_R8;V_R8(&variants[5])=2.75;
    V_VT(&variants[6])=VT_CY;V_CY(&variants[6]).int64=123456;
    V_VT(&variants[7])=VT_DATE;V_DATE(&variants[7])=45123.25;
    V_VT(&variants[8])=VT_DECIMAL;V_DECIMAL(&variants[8]).scale=3;V_DECIMAL(&variants[8]).Lo64=1125;
    for(int i=0;i<9;i++){VARIANT before=variants[i];Oper *result=variant_echo(&variants[i]);CHECK(result && result->type==(1|16384) && result->value.number==expected[i]);CHECK(memcmp(&before,&variants[i],sizeof(before))==0);release(result);CHECK(owned()==0 && count()==0);}
    Oper input={0},*output=NULL;
    input.type=1;input.value.number=123.5;
    output=echo(&input);CHECK(output && output->type==(1|16384) && output->value.number==123.5);release(output); CHECK(owned()==0 && count()==0);
    WCHAR text[]={4,L'A',0,0x3bb,0x3a9};
    input.type=2;input.value.text=text;
    output=echo(&input);CHECK(output && output->type==(2|16384));CHECK(output->value.text!=text && memcmp(output->value.text,text,sizeof(text))==0);release(output);CHECK(owned()==0 && count()==0);
    input.type=4;input.value.integer=-1;
    output=echo(&input);CHECK(output && output->type==(4|16384) && output->value.integer==1);release(output);CHECK(owned()==0);
    LONG errors[]={0,7,15,23,29,36,42,43};
    for(int i=0;i<8;i++){input.type=16;input.value.integer=errors[i];output=echo(&input);if(!output || output->type!=(16|16384) || output->value.integer!=errors[i])fprintf(stderr,"error %ld -> type=%ld value=%ld VBA=%ld\n",errors[i],output?output->type:-1,output?output->value.integer:-1,error(NULL,NULL));CHECK(output && output->type==(16|16384) && output->value.integer==errors[i]);release(output);CHECK(owned()==0);}
    input.type=128;output=echo(&input);CHECK(output && output->type==(128|16384));release(output);CHECK(owned()==0);
    input.type=256;output=echo(&input);CHECK(output && output->type==(256|16384));release(output);CHECK(owned()==0);
    input.type=2048;input.value.integer=-3;output=echo(&input);CHECK(output && output->type==(1|16384) && output->value.number==-3);release(output);CHECK(owned()==0);
    Oper cells[6]={0};
    cells[0].type=1;cells[0].value.number=2.5;cells[1].type=2;cells[1].value.text=text;cells[2].type=4;cells[2].value.integer=1;
    cells[3].type=16;cells[3].value.integer=42;cells[4].type=128;cells[5].type=256;
    input.type=64;input.value.multi.values=cells;input.value.multi.rows=2;input.value.multi.columns=3;
    output=echo(&input);CHECK(output && output->type==(64|16384) && output->value.multi.rows==2 && output->value.multi.columns==3);
    for(int i=0;i<6;i++){CHECK(output->value.multi.values[i].type==cells[i].type);}
    CHECK(output->value.multi.values[0].value.number==2.5 && output->value.multi.values[3].value.integer==42);
    CHECK(memcmp(output->value.multi.values[1].value.text,text,sizeof(text))==0);release(output);CHECK(owned()==0);
    input.type=1024;input.value.sref.count=1;input.value.sref.padding=0;input.value.sref.area=(Rect){2,4,1,3};
    output=echo(&input);CHECK(output && output->type==(1024|16384) && output->value.sref.count==1 && memcmp(&output->value.sref.area,&input.value.sref.area,16)==0);release(output);CHECK(owned()==0);
    struct {WORD count,padding;Rect area[2];} areas={2,0,{{0,1,2,3},{5,6,7,8}}};
    input.type=8;input.value.ref.areas=&areas;input.value.ref.sheet=0x123456789abcdefLL;
    output=echo(&input);CHECK(output && output->type==(8|16384) && output->value.ref.sheet==input.value.ref.sheet);
    CHECK(output->value.ref.areas!=&areas && memcmp(output->value.ref.areas,&areas,sizeof(areas))==0);release(output);CHECK(owned()==0);
    WCHAR *large=HeapAlloc(GetProcessHeap(),0,32768*sizeof(WCHAR));CHECK(large);large[0]=32767;
    for(int i=1;i<32768;i++)large[i]=(WCHAR)(0x300+(i%200));
    input.type=2;input.value.text=large;output=echo(&input);CHECK(output && output->type==(2|16384) && memcmp(output->value.text,large,32768*sizeof(WCHAR))==0);release(output);HeapFree(GetProcessHeap(),0,large);CHECK(owned()==0 && count()==0);
    printf("numeric-types|number|string|boolean|errors|missing|nil|integer|matrix|sref|mref|long-string|owned=0\n");
    if(argc==3){LARGE_INTEGER frequency,start,end;QueryPerformanceFrequency(&frequency);input.type=1;input.value.number=2.5;QueryPerformanceCounter(&start);
    for(int i=0;i<10000;i++){output=echo(&input);CHECK(output);release(output);}QueryPerformanceCounter(&end);
    printf("Echo_and_free_us=%.3f owned=%ld\n",(double)(end.QuadPart-start.QuadPart)*1e6/(double)frequency.QuadPart/10000,owned());}
    Count shutdown=(Count)GetProcAddress(dll,"CalcVbaShutdown");CHECK(shutdown && shutdown()==0);CHECK(FreeLibrary(dll));return 0;
}
