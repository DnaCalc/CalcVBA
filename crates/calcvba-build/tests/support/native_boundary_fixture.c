#define COBJMACROS
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oleauto.h>
#include <stdint.h>

/* Registration-free COM identity fixture. Two distinct interface addresses
   return one canonical IUnknown; a third belongs to a different identity. */
typedef struct Projection {
    IUnknown iface;
    struct Projection *identity;
    LONG references;
} Projection;
static HRESULT STDMETHODCALLTYPE query(IUnknown *self, REFIID iid, void **out) {
    Projection *value = (Projection *)self;
    (void)iid;
    *out = &value->identity->iface;
    IUnknown_AddRef((IUnknown *)*out);
    return S_OK;
}
static ULONG STDMETHODCALLTYPE addref(IUnknown *self) {
    Projection *value = (Projection *)self;
    return (ULONG)InterlockedIncrement(&value->identity->references);
}
static ULONG STDMETHODCALLTYPE release(IUnknown *self) {
    Projection *value = (Projection *)self;
    return (ULONG)InterlockedDecrement(&value->identity->references);
}
static IUnknownVtbl table = {query, addref, release};
static Projection first = {{&table}, &first, 0};
static Projection alternate = {{&table}, &first, 0};
static Projection second = {{&table}, &second, 0};
__declspec(dllexport) void *FixtureProjection(LONG index) {
    return index == 0 ? &first.iface : index == 1 ? &alternate.iface : &second.iface;
}
__declspec(dllexport) LONG FixtureReferences(void) {
    return first.references + second.references;
}

#include <string.h>
#include <wchar.h>
#include <stdio.h>
__declspec(dllexport) BSTR FixtureArrayShape(SAFEARRAY **cell) {
    char text[128]; SAFEARRAY *a=*cell; VARTYPE vt=0; IRecordInfo *info=NULL;
    HRESULT kind=SafeArrayGetVartype(a,&vt), record=SafeArrayGetRecordInfo(a,&info);
    if(info) IRecordInfo_Release(info);
    int n=sprintf_s(text,sizeof(text),"%u:%u:%ld:%u:%ld",a->fFeatures,a->cbElements,(LONG)kind,vt,(LONG)record);
    return SysAllocStringByteLen(text,(UINT)n);
}
typedef struct Payload { BYTE tag; char fixed[4]; LONG number; BSTR text; VARIANT value; } Payload;
typedef struct Pair { LONG first, second; } Pair;
typedef struct Returned { LONG first, second; BSTR text; VARIANT value; } Returned;
__declspec(dllexport) LONG FixtureRecord(Payload *p) {
    if (p->tag!=7 || memcmp(p->fixed,"ABCD",4)!=0 || p->number!=37 || V_VT(&p->value)!=VT_BSTR) return -1;
    if (SysStringByteLen(p->text)!=3 || memcmp(p->text,"abc",3)!=0 || wcscmp(V_BSTR(&p->value),L"wide")!=0) return -2;
    memcpy(p->fixed,"wxyz",4); p->number=42;
    SysFreeString(p->text); p->text=SysAllocStringByteLen("new",3);
    VariantClear(&p->value); V_VT(&p->value)=VT_BSTR; V_BSTR(&p->value)=SysAllocString(L"changed");
    return 1;
}
__declspec(dllexport) Pair FixturePair(LONG value) { Pair p={value,value+7}; return p; }
__declspec(dllexport) Returned FixtureReturned(void) {
    Returned r={0}; r.first=6; r.second=9; r.text=SysAllocStringByteLen("result",6);
    V_VT(&r.value)=VT_BSTR; V_BSTR(&r.value)=SysAllocString(L"wide"); return r;
}
__declspec(dllexport) VARIANT FixtureVariant(LONG value) { VARIANT v; VariantInit(&v); V_VT(&v)=VT_I4; V_I4(&v)=value+5; return v; }
__declspec(dllexport) VARIANT FixtureVariantText(void) { VARIANT v; VariantInit(&v); V_VT(&v)=VT_BSTR; V_BSTR(&v)=SysAllocString(L"variant"); return v; }
__declspec(dllexport) LONG FixtureAnsiAlias(BSTR *a,BSTR *b) {
    LONG distinct=a!=b;
    SysFreeString(*a); *a=SysAllocStringByteLen("first",5);
    SysFreeString(*b); *b=SysAllocStringByteLen("last",4);
    return distinct;
}
