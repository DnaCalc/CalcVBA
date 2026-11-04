#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oleauto.h>
#include <stdint.h>
#include <inttypes.h>
#include <stdio.h>
#include <wchar.h>
#define CHECK(x) do { if(!(x)) { fprintf(stderr,"line %d: %s\n",__LINE__,#x); return 2; } } while(0)
#define LOAD(type,name) type name=(type)GetProcAddress(dll,#name); CHECK(name)
typedef uint8_t (*byte_ref)(uint8_t*);
typedef int16_t (*integer_ref)(int16_t*);
typedef int64_t (*wide_ref)(int64_t*);
typedef float (*single_ref)(float*);
typedef double (*date_ref)(double*);
typedef double (*sum_widths)(uint8_t,int16_t,int64_t,float,int64_t,double);
typedef BSTR (*string_zero)(void);
typedef int32_t (*error_get)(BSTR*,BSTR*);
static int first=1;
static void number(int64_t value) { printf("%s%" PRId64,first?"":"|",value); first=0; }
static void real(double value) { printf("%s%.17g",first?"":"|",value); first=0; }
int wmain(int argc,wchar_t**argv) {
    CHECK(argc==2); HMODULE dll=LoadLibraryW(argv[1]); CHECK(dll);
    LOAD(byte_ref,IncByte); LOAD(integer_ref,IncInteger); LOAD(wide_ref,IncLongLong);
    LOAD(wide_ref,IncPointer); LOAD(single_ref,IncSingle); LOAD(wide_ref,IncCurrency);
    LOAD(date_ref,IncDate); LOAD(sum_widths,SumWidths); LOAD(string_zero,WideSuite);
    LOAD(error_get,CalcVbaGetLastError);
    struct { uint8_t value,guard; } b={254,0xa5};
    struct { int16_t value,guard; } i={-2,0x4567};
    int64_t ll=INT64_C(9007199254740993),p=INT64_C(81985529216486895),c=12345;
    float s=1.25f; double d=45000.25;
    number(IncByte(&b.value)); number(b.value); CHECK(b.guard==0xa5);
    number(IncInteger(&i.value)); number(i.value); CHECK(i.guard==0x4567);
    number(IncLongLong(&ll)); number(ll);
    number(IncPointer(&p)); number(p);
    real(IncSingle(&s)); real(s);
    number(IncCurrency(&c)); number(c);
    real(IncDate(&d)); real(d);
    real(SumWidths(5,-3,100,1.25f,25000,45000.75));
    BSTR text=WideSuite(); CHECK(text);
    CHECK(wcscmp(text,L"255:255:-1:-1:9007199254740994:9007199254740994:81985529216486903:81985529216486903:1.75:1.75:12346:12346:45000.75:45000.75:45106.5")==0);
    number(SysStringLen(text)); SysFreeString(text);
    number(IncByte(&b.value)); number(b.value); number(CalcVbaGetLastError(NULL,NULL)); CHECK(b.guard==0xa5);
    i.value=32767; number(IncInteger(&i.value)); number(i.value); number(CalcVbaGetLastError(NULL,NULL)); CHECK(i.guard==0x4567);
    ll=INT64_MAX; number(IncLongLong(&ll)); number(ll); number(CalcVbaGetLastError(NULL,NULL));
    puts(""); CHECK(FreeLibrary(dll)); return 0;
}
