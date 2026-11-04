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
