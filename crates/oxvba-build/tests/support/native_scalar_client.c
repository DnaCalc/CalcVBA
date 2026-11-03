#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdint.h>
#include <stdio.h>
#include <math.h>

typedef int32_t (*long2)(int32_t, int32_t);
typedef double (*double2)(double, double);
typedef double (*double1)(double);
typedef int16_t (*compare)(int32_t, int32_t);
typedef int16_t (*boolean1)(int16_t);
typedef double (*mixed5)(double, int32_t, int16_t, double, int32_t);
struct wrong_thread { long2 call; int32_t result; DWORD error; };
static DWORD WINAPI call_wrong_thread(void *pointer) {
    struct wrong_thread *probe = pointer;
    probe->result = probe->call(7,35);
    probe->error = GetLastError();
    return 0;
}

int wmain(int argc, wchar_t **argv) {
    if (argc != 2) return 2;
    HMODULE dll = LoadLibraryW(argv[1]);
    if (!dll) { fprintf(stderr,"LoadLibrary: %lu\n",GetLastError()); return 3; }
    long2 add = (long2)GetProcAddress(dll,"AddLong");
    long2 direct = (long2)GetProcAddress(dll,"DirectLong");
    double2 add_double = (double2)GetProcAddress(dll,"AddDouble");
    compare greater = (compare)GetProcAddress(dll,"Greater");
    boolean1 echo = (boolean1)GetProcAddress(dll,"EchoBoolean");
    mixed5 mixed = (mixed5)GetProcAddress(dll,"Mixed5");
    double2 divide = (double2)GetProcAddress(dll,"Divide");
    double1 library_abs = (double1)GetProcAddress(dll,"LibraryAbs");
    double1 library_sqr = (double1)GetProcAddress(dll,"LibrarySqr");
    if (!add || !direct || !add_double || !greater || !echo || !mixed || !divide || !library_abs || !library_sqr) return 4;
    if ((FARPROC)add != GetProcAddress(dll,MAKEINTRESOURCEA(11))) return 5;
    printf("%d|%d|%.17g|%d|%d|%d|%d|%.17g|%.17g|%.17g|%.17g|%d|%.17g|%.17g|%.17g\n",
        add(-7,49),direct(5,32),add_double(1.25,2.5),
        greater(5,-2),greater(-2,5),echo(-1),echo(0),
        mixed(1.25,2,-1,3.5,4),mixed(1.25,2,0,3.5,4),
        divide(9,2),divide(9,0),add(INT32_MAX,1),library_abs(-12.5),library_sqr(25),library_sqr(-1));
    if (add_double(NAN,1) != 0 || library_abs(INFINITY) != 0) return 6;
    struct wrong_thread probe = {add, -1, 0};
    HANDLE thread = CreateThread(NULL,0,call_wrong_thread,&probe,0,NULL);
    if (!thread) return 7;
    if (WaitForSingleObject(thread,5000) != WAIT_OBJECT_0) return 8;
    CloseHandle(thread);
    if (probe.result != 0 || probe.error != 1444) return 9;
    FreeLibrary(dll);
    return 0;
}
