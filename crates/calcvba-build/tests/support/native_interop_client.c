#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
#include <windows.h>
#include <oleauto.h>
#include <stdint.h>
#include <stdio.h>
#include <stddef.h>
#include <wchar.h>

#define CHECK(test) do { if (!(test)) { fprintf(stderr,"line %d: %s (Win32 %lu)\n",__LINE__,#test,GetLastError()); return 2; } } while (0)
#define LOAD(type, name) type name = (type)GetProcAddress(dll,#name); CHECK(name)
typedef int32_t (*long_refs)(int32_t*,int32_t*);
typedef double (*double_ref)(double*);
typedef int16_t (*bool_ref)(int16_t*);
typedef int32_t (*long_value)(int32_t);
typedef int32_t (*long_zero)(void);
typedef int16_t (*string_pair)(BSTR,BSTR);
typedef BSTR (*array_string)(SAFEARRAY**);
typedef BSTR (*string_zero)(void);
typedef struct SmallNativeRecord { int32_t value; } SmallNativeRecord;
typedef SmallNativeRecord (*small_record)(int32_t);
typedef int32_t (*long_ref)(int32_t*);
typedef void (*sub_ref)(int32_t*);
typedef double (*double_pair)(double,double);
typedef int32_t (*double_hr)(double,double,double*);
typedef int32_t (*last_error)(BSTR*,BSTR*);
typedef BSTR (*string_value)(BSTR);
typedef int32_t (*string_ref)(BSTR*);
typedef int32_t (*string_length)(const void*);
typedef void* (*variant_result)(VARIANT*,VARIANT*,int32_t);
typedef BSTR (*variant_string)(const VARIANT*);
typedef int32_t (*variant_long)(const VARIANT*);
typedef IUnknown* (*object_value)(IUnknown*);
typedef int16_t (*object_bool)(IUnknown*);
typedef int16_t (*object_pair)(IUnknown*,IUnknown*);
typedef void (*object_ref)(IUnknown**);
typedef struct NativeRecord {
    uint8_t tag;
    int16_t code;
    int32_t count;
    double amount;
    int16_t enabled;
    BSTR text;
} NativeRecord;
typedef double (*record_change)(NativeRecord*);
typedef void* (*record_result)(NativeRecord*,const NativeRecord*);
typedef int32_t (*array_long)(SAFEARRAY**);
typedef void (*array_allocate)(SAFEARRAY**);
typedef double (*array_matrix)(SAFEARRAY**,int32_t,int32_t);
typedef double (*array_double)(SAFEARRAY**);

typedef struct Mock { IUnknown object; LONG references; } Mock;
static HRESULT STDMETHODCALLTYPE query(IUnknown* object, REFIID iid, void** out) {
    (void)iid;
    *out=object;
    IUnknown_AddRef(object);
    return S_OK;
}
static ULONG STDMETHODCALLTYPE addref(IUnknown* object) { return (ULONG)InterlockedIncrement(&((Mock*)object)->references); }
static ULONG STDMETHODCALLTYPE release(IUnknown* object) { return (ULONG)InterlockedDecrement(&((Mock*)object)->references); }
static IUnknownVtbl table = {query,addref,release};
static int first=1;
static void integer(int32_t value) { printf("%s%d",first?"":"|",value); first=0; }
static void floating(double value) { printf("%s%.17g",first?"":"|",value); first=0; }

struct ThreadProbe {
    long_value next;
    long_refs bump;
    string_ref mutate;
    double_hr hr;
    last_error error;
    int passed;
};
static DWORD WINAPI wrong_thread(void* state) {
    struct ThreadProbe* p=state;
    int32_t value=17;
    BSTR string=SysAllocString(L"owned");
    BSTR original=string;
    double result=77.5;
    BSTR source=(BSTR)(uintptr_t)1, description=(BSTR)(uintptr_t)2;
    int passed=p->next(1000)==0 && GetLastError()==1444;
    passed=passed && p->bump(&value,&value)==0 && value==17;
    passed=passed && p->mutate(&string)==0 && string==original && SysStringLen(string)==5;
    passed=passed && (uint32_t)p->hr(9,2,&result)==0x8001010eu && result==77.5;
    passed=passed && p->error(&source,&description)==1444 && source==(BSTR)(uintptr_t)1 && description==(BSTR)(uintptr_t)2;
    SysFreeString(string);
    p->passed=passed;
    return 0;
}

int wmain(int argc, wchar_t** argv) {
    CHECK(argc==2);
    HMODULE dll=LoadLibraryW(argv[1]); CHECK(dll);
    LOAD(long_value,RaiseCustom); LOAD(long_zero,ReadLastDllError); LOAD(string_pair,SameString); LOAD(array_string,JoinArray);
    LOAD(long_refs,Bump); LOAD(double_ref,BumpDouble); LOAD(bool_ref,Flip);
    LOAD(long_value,NextValue); LOAD(long_value,StaticValue);
    LOAD(sub_ref,Mutate); LOAD(long_ref,Nested);
    LOAD(double_pair,Divide); LOAD(double_hr,DivideHResult);
    LOAD(last_error,CalcVbaGetLastError);
    LOAD(string_value,EchoString); LOAD(string_ref,MutateString);
    LOAD(string_length,StringLength); LOAD(string_length,LengthAnsi); LOAD(string_length,LengthUtf16);
    LOAD(variant_result,AddVariant); LOAD(variant_string,VariantText);
    LOAD(object_value,EchoObject); LOAD(object_pair,SameObject); LOAD(object_ref,ClearObject);
    LOAD(record_change,ChangeRecord); LOAD(record_result,EchoRecord);
    LOAD(array_long,SumArray); LOAD(array_long,ChangeArray); LOAD(array_allocate,AllocateArray); LOAD(array_matrix,ReadMatrix);
    typedef char record_layout_assert[(sizeof(NativeRecord)==32 && offsetof(NativeRecord,text)==24)?1:-1];
    (void)sizeof(record_layout_assert);
    int32_t a=5,b=7;
    integer(Bump(&a,&b)); integer(a); integer(b);
    a=5; integer(Bump(&a,&a)); integer(a);
    double d=1.25; floating(BumpDouble(&d)); floating(d);
    int16_t boolean=-1; integer(Flip(&boolean)); integer(Flip(&boolean));
    integer(NextValue(3)); integer(NextValue(7)); integer(StaticValue(4)); integer(StaticValue(5));
    a=5; Mutate(&a); integer(a); a=4; integer(Nested(&a)); integer(a);
    floating(Divide(9,2)); floating(Divide(9,0));
    BSTR source=NULL,description=NULL;
    integer(CalcVbaGetLastError(&source,&description));
    CHECK(wcscmp(source,L"NativeInterop")==0 && wcscmp(description,L"Division by zero")==0);
    SysFreeString(source); SysFreeString(description);
    d=99; integer(DivideHResult(9,0,&d)); floating(d);
    integer(DivideHResult(9,2,&d)); floating(d);
    CHECK(CalcVbaGetLastError(NULL,NULL)==0);
    CHECK(Bump(NULL,&a)==0 && CalcVbaGetLastError(NULL,NULL)==5);
    wchar_t units[]={L'A',0,0x03a9};
    BSTR text=SysAllocStringLen(units,3); CHECK(text);
    BSTR result=EchoString(text); CHECK(result);
    integer((int32_t)SysStringLen(result)); CHECK(result[0]==L'<' && result[2]==0 && result[3]==0x03a9 && result[4]==L'>');
    CHECK(SysStringLen(text)==3); SysFreeString(result);
    integer(StringLength(text)); integer(MutateString(&text)); CHECK(text[3]==L'!'); SysFreeString(text);
    integer(LengthAnsi("Hello")); integer(LengthUtf16(L"\x6f22\x5b57x"));
    VARIANT value={0},returned={0}; V_VT(&value)=VT_I4; V_I4(&value)=37;
    CHECK(AddVariant(&returned,&value,5)==&returned);
    integer(V_I4(&value)); integer(V_I4(&returned)); CHECK(V_VT(&value)==VT_I4 && V_VT(&returned)==VT_I4);
    VariantClear(&returned);
    result=VariantText(&value); CHECK(result && wcscmp(result,L"42")==0); integer((int32_t)SysStringLen(result)); SysFreeString(result);
    VariantClear(&value);
    Mock object={{&table},1};
    IUnknown* echoed=EchoObject(&object.object); CHECK(echoed==&object.object && object.references==2);
    integer(SameObject(&object.object,echoed)); integer(SameObject(&object.object,NULL));
    ClearObject(&echoed); CHECK(!echoed && object.references==1);
    NativeRecord record={0}; record.tag=127; record.code=-123; record.count=5; record.amount=1.25; record.enabled=-1; record.text=SysAllocString(L"ab");
    floating(ChangeRecord(&record)); integer(record.count); floating(record.amount); integer(record.enabled); integer((int32_t)SysStringLen(record.text));
    CHECK(record.tag==127 && record.code==-123 && wcscmp(record.text,L"ab!")==0);
    NativeRecord copy={0}; CHECK(EchoRecord(&copy,&record)==&copy);
    CHECK(copy.text!=record.text && wcscmp(copy.text,record.text)==0); integer(copy.count);
    SysFreeString(copy.text); SysFreeString(record.text);
    SAFEARRAY* array=NULL; AllocateArray(&array); CHECK(array);
    integer(SumArray(&array)); integer(ChangeArray(&array));
    LONG lower=0,upper=0; CHECK(SUCCEEDED(SafeArrayGetLBound(array,1,&lower)) && SUCCEEDED(SafeArrayGetUBound(array,1,&upper)) && lower==-2 && upper==1);
    SafeArrayDestroy(array);
    SAFEARRAYBOUND bounds[2]={{2,-1},{3,3}};
    SAFEARRAY* matrix=SafeArrayCreate(VT_R8,2,bounds); CHECK(matrix);
    LONG index[2]={0,5}; d=3.5; CHECK(SUCCEEDED(SafeArrayPutElement(matrix,index,&d)));
    floating(ReadMatrix(&matrix,0,5)); CHECK(SUCCEEDED(SafeArrayGetElement(matrix,index,&d))); floating(d);
    LOAD(array_double,GrowMatrix); floating(GrowMatrix(&matrix)); index[1]=6;
    CHECK(SUCCEEDED(SafeArrayGetElement(matrix,index,&d))); floating(d);
    CHECK(SUCCEEDED(SafeArrayGetUBound(matrix,2,&upper))); integer(upper); SafeArrayDestroy(matrix);
    integer(RaiseCustom(513)); source=NULL; description=NULL;
    integer(CalcVbaGetLastError(&source,&description));
    CHECK(wcscmp(source,L"NativeCase")==0 && wcscmp(description,L"custom failure")==0);
    SysFreeString(source); SysFreeString(description);
    SetLastError(12345); int32_t observed=ReadLastDllError(); CHECK(observed==12345 && GetLastError()==12345); integer(observed);
    BSTR left=SysAllocStringLen(units,3); units[2]=0x03aa; BSTR right=SysAllocStringLen(units,3);
    integer(SameString(left,left)); integer(SameString(left,right)); SysFreeString(left); SysFreeString(right);
    SAFEARRAY* strings=SafeArrayCreateVector(VT_BSTR,-2,2); CHECK(strings);
    LONG first_index=-2,last_index=-1; left=SysAllocString(L"left"); right=SysAllocString(L"right");
    CHECK(SUCCEEDED(SafeArrayPutElement(strings,&first_index,left)) && SUCCEEDED(SafeArrayPutElement(strings,&last_index,right)));
    SysFreeString(left); SysFreeString(right);
    result=JoinArray(&strings); CHECK(wcscmp(result,L"left|right")==0); integer((int32_t)SysStringLen(result)); SysFreeString(result);
    CHECK(SUCCEEDED(SafeArrayGetElement(strings,&last_index,&result))); integer((int32_t)SysStringLen(result)); CHECK(wcscmp(result,L"right!")==0); SysFreeString(result); SafeArrayDestroy(strings);
    V_VT(&value)=VT_BSTR; V_BSTR(&value)=SysAllocString(L"37");
    AddVariant(&returned,&value,5); integer(V_VT(&value)); integer(V_VT(&returned)); floating(V_R8(&returned)); VariantClear(&value); VariantClear(&returned);
    V_VT(&value)=VT_BOOL; V_BOOL(&value)=-1;
    result=VariantText(&value); CHECK(wcscmp(result,L"True")==0); integer((int32_t)SysStringLen(result)); SysFreeString(result); VariantClear(&value);
    V_VT(&value)=VT_NULL; AddVariant(&returned,&value,5); integer(V_VT(&value)); integer(V_VT(&returned)); VariantClear(&returned);
    result=VariantText(&value); integer(result==NULL?0:1); CHECK(result==NULL); integer(CalcVbaGetLastError(NULL,NULL)); VariantClear(&value);
    wchar_t lone=0xd800; text=SysAllocStringLen(&lone,1); result=EchoString(text); CHECK(SysStringLen(result)==3 && result[1]==0xd800);
    integer((int32_t)SysStringLen(result)); integer((int16_t)result[1]); SysFreeString(text); SysFreeString(result);
    struct ThreadProbe probe={NextValue,Bump,MutateString,DivideHResult,CalcVbaGetLastError,0};
    HANDLE thread=CreateThread(NULL,0,wrong_thread,&probe,0,NULL); CHECK(thread);
    CHECK(WaitForSingleObject(thread,10000)==WAIT_OBJECT_0); CloseHandle(thread); CHECK(probe.passed);
    CHECK(CalcVbaGetLastError(NULL,NULL)==1444 && NextValue(0)==10 && CalcVbaGetLastError(NULL,NULL)==0);
    LOAD(small_record,SmallRecord); integer(SmallRecord(37).value);
    LOAD(string_zero,ValueSuite); result=ValueSuite();
    CHECK(result && wcscmp(result,L"19:6:13:24:12:3.5:1.75:False:<omega>:omega!:6:42:42:8.75:ab!:49:-2:3.5:3.75:3.75:6:7.5:left|right:right!:True:9:9:7")==0);
    integer((int32_t)SysStringLen(result)); SysFreeString(result);
    LOAD(variant_long,VariantTruth); VariantInit(&value); integer(VariantTruth(&value));
    V_VT(&value)=VT_NULL; integer(VariantTruth(&value));
    V_VT(&value)=VT_I4; V_I4(&value)=0; integer(VariantTruth(&value));
    V_I4(&value)=1; integer(VariantTruth(&value)); VariantClear(&value);
    LOAD(string_pair,TextSame); left=SysAllocString(L"\x00c5"); right=SysAllocString(L"\x00e5");
    integer(TextSame(left,right)); SysFreeString(left); SysFreeString(right);
    LOAD(string_value,StoreString); left=SysAllocString(L"a"); right=SysAllocString(L"b");
    result=StoreString(left); CHECK(wcscmp(result,L"a")==0); integer((int32_t)SysStringLen(result)); SysFreeString(result);
    result=StoreString(right); CHECK(wcscmp(result,L"ab")==0); integer((int32_t)SysStringLen(result)); SysFreeString(result);
    SysFreeString(left); SysFreeString(right);
    LOAD(object_bool,HoldObject); integer(HoldObject(&object.object)); CHECK(object.references==2);
    LOAD(long_zero,CalcVbaShutdown); integer(CalcVbaShutdown()); CHECK(object.references==1);
    integer(NextValue(0)); integer(CalcVbaGetLastError(NULL,NULL));
    printf("\n");
    CHECK(FreeLibrary(dll));
    return 0;
}
