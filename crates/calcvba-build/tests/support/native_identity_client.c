#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdio.h>
typedef void *(*Projection)(LONG);
typedef LONG (*Count)(void);
typedef SHORT (*Same)(void *, void *);
int wmain(int argc, wchar_t **argv) {
    if (argc != 3) return 2;
    HMODULE native = LoadLibraryW(argv[1]), fixture = LoadLibraryW(argv[2]);
    if (!native || !fixture) return 3;
    Projection projection = (Projection)GetProcAddress(fixture, "FixtureProjection");
    Count references = (Count)GetProcAddress(fixture, "FixtureReferences");
    Same same = (Same)GetProcAddress(native, "SameObject");
    Count shutdown = (Count)GetProcAddress(native, "CalcVbaShutdown");
    if (!projection || !references || !same || !shutdown) return 4;
    void *first = projection(0), *alternate = projection(1), *second = projection(2);
    if (first == alternate || references() != 0) return 5;
    SHORT a = same(first, alternate), b = same(first, second), c = same(NULL, NULL);
    LONG balance = references();
    if (a != -1 || b != 0 || c != -1 || balance != 0 || shutdown() != 0) return 6;
    printf("%d|%d|%d|%ld\n", a, b, c, balance);
    FreeLibrary(native); FreeLibrary(fixture);
    return 0;
}
