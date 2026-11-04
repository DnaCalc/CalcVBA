#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <stddef.h>
#include <wchar.h>
#include <string.h>
#define CHECK(test) do { if (!(test)) { fprintf(stderr,"line %d: %s\n", __LINE__, #test); return 2; } } while(0)
typedef struct RichRecord { BYTE tag; VARIANT value; LONG tail; WCHAR text[4]; } RichRecord;
typedef BSTR (*Change)(RichRecord *);
typedef void *(*Echo)(RichRecord *, const RichRecord *);
typedef BSTR (*Suite)(void);
typedef LONG (*Grow)(SAFEARRAY **);
typedef void (*Allocate)(SAFEARRAY **);
typedef LONG (*Shutdown)(void);
typedef BSTR (*RecordArrays)(SAFEARRAY**);
typedef LONG (*Clear)(BSTR*);
int wmain(int argc, wchar_t **argv) {
    setvbuf(stdout,NULL,_IONBF,0);
    CHECK(argc == 2 && sizeof(RichRecord) == 48 && offsetof(RichRecord,value) == 8 && offsetof(RichRecord,tail) == 32 && offsetof(RichRecord,text) == 36);
    HMODULE dll = LoadLibraryW(argv[1]); CHECK(dll);
    Change change = (Change)GetProcAddress(dll,"ChangeRecord");
    Echo echo = (Echo)GetProcAddress(dll,"EchoRecord");
    Suite suite = (Suite)GetProcAddress(dll,"ValueSuite");
    Grow grow = (Grow)GetProcAddress(dll,"GrowStrings");
    Allocate allocate = (Allocate)GetProcAddress(dll,"AllocateStrings");
    Allocate ansi_allocate = (Allocate)GetProcAddress(dll,"AllocateAnsiStrings");
    Shutdown shutdown = (Shutdown)GetProcAddress(dll,"CalcVbaShutdown");
    Clear clear=(Clear)GetProcAddress(dll,"ClearAnsi"); Suite ansi_text=(Suite)GetProcAddress(dll,"AnsiText");
    CHECK(change && echo && suite && grow && allocate && ansi_allocate && shutdown && clear && ansi_text);
    RichRecord record = {0}, copied = {0}; record.tag=7; record.tail=9;
    V_VT(&record.value)=VT_BSTR; V_BSTR(&record.value)=SysAllocString(L"7");
    memcpy(record.text,L"q   ",8);
    CHECK(echo(&copied,&record)==&copied && V_BSTR(&copied.value)!=V_BSTR(&record.value));
    BSTR description=change(&record);
    if (!description || wcscmp(description,L"12:abcd:10")!=0) {
        typedef LONG (*Error)(BSTR*,BSTR*); Error error=(Error)GetProcAddress(dll,"CalcVbaGetLastError");
        BSTR source=NULL,message=NULL; LONG code=error(&source,&message);
        fwprintf(stderr,L"ChangeRecord: code=%ld text=%s source=%s message=%s vt=%u tail=%ld fixed=%.4s\n",code,description?description:L"<null>",source?source:L"",message?message:L"",V_VT(&record.value),record.tail,record.text);
        SysFreeString(source); SysFreeString(message);
    }
    CHECK(description && wcscmp(description,L"12:abcd:10")==0 && V_VT(&record.value)==VT_R8 && V_R8(&record.value)==12 && record.tail==10 && record.tag==7 && memcmp(record.text,L"abcd",8)==0);
    CHECK(wcscmp(V_BSTR(&copied.value),L"7")==0 && memcmp(copied.text,L"q   ",8)==0);
    BSTR result=suite(); CHECK(result && wcscmp(result,L"12:abcd:10:12:abcd:10")==0);
    printf("12:abcd:10|7|4|%u|",SysStringLen(result));
    SysFreeString(description); SysFreeString(result); VariantClear(&record.value); VariantClear(&copied.value);
    SAFEARRAY *values=NULL; ansi_allocate(&values); CHECK(values);
    LONG count=grow(&values), lower=0, upper=0, index=0; VARTYPE kind=99; BSTR text=NULL;
    CHECK(SafeArrayGetVartype(values,&kind)==S_OK && kind==VT_EMPTY && (values->fFeatures&FADF_BSTR));
    CHECK(SafeArrayGetLBound(values,1,&lower)==S_OK && SafeArrayGetUBound(values,1,&upper)==S_OK);
    CHECK(SafeArrayGetElement(values,&index,&text)==S_OK && SysStringByteLen(text)==3 && memcmp(text,"new",3)==0);
    printf("%ld|%ld|%ld|%u|",count,lower,upper,SysStringByteLen(text)); SysFreeString(text); CHECK(SafeArrayDestroy(values)==S_OK);
    values=NULL; allocate(&values); CHECK(values); CHECK(SafeArrayGetVartype(values,&kind)==S_OK && kind==VT_BSTR);
    count=grow(&values); CHECK(SafeArrayGetElement(values,&index,&text)==S_OK && wcscmp(text,L"new")==0); SysFreeString(text);
    printf("%ld|",count);
    BSTR ansi=SysAllocString(L"initial"); LONG cleared=clear(&ansi); CHECK(cleared==1 && SysStringLen(ansi)==0); SysFreeString(ansi);
    ansi=ansi_text(); CHECK(ansi && wcscmp(ansi,L"reply")==0); printf("%ld|0|%u|",cleared,SysStringLen(ansi));
    Allocate allocate_records=(Allocate)GetProcAddress(dll,"AllocateRecords"); RecordArrays record_arrays=(RecordArrays)GetProcAddress(dll,"RecordArrays"); Suite array_suite=(Suite)GetProcAddress(dll,"ArraySuite"); CHECK(allocate_records && record_arrays && array_suite);
    SAFEARRAY *records=NULL; allocate_records(&records); CHECK(records && records->cbElements==48); BSTR array_result=record_arrays(&records); CHECK(array_result && wcscmp(array_result,L"7:q   :-2:0")==0); SysFreeString(array_result);
    array_result=array_suite(); CHECK(array_result && wcscmp(array_result,L"7:q   :-2:0")==0); SysFreeString(array_result); CHECK(SafeArrayDestroy(records)==S_OK);
    Suite fixed_suite=(Suite)GetProcAddress(dll,"FixedFieldSuite"); CHECK(fixed_suite);
    BSTR fixed_result=fixed_suite();
    if(!fixed_result || wcscmp(fixed_result,L"29:41:left:42:abc:q  :record")!=0) fwprintf(stderr,L"FixedFieldSuite: %s\n",fixed_result?fixed_result:L"<null>");
    CHECK(fixed_result && wcscmp(fixed_result,L"29:41:left:42:abc:q  :record")==0); SysFreeString(fixed_result);
    typedef LONG (*Index)(LONG); typedef LONG (*Error)(BSTR*,BSTR*);
    Index fixed_index=(Index)GetProcAddress(dll,"FixedFieldIndex"); Error error=(Error)GetProcAddress(dll,"CalcVbaGetLastError");
    CHECK(fixed_index && error && fixed_index(-1)==29);
    CHECK(fixed_index(-2)==0 && error(NULL,NULL)==9); CHECK(fixed_index(2)==0 && error(NULL,NULL)==9);
    printf("7:q   :-2:0|7:q   :-2:0|%ld\n",shutdown()); SysFreeString(ansi); CHECK(SafeArrayDestroy(values)==S_OK);
    CHECK(FreeLibrary(dll)); return 0;
}
