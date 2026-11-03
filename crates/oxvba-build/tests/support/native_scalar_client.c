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
typedef int16_t (*boolean2)(int16_t,int16_t);
typedef int32_t (*long1)(int32_t);
typedef int32_t (*double_long)(double);
typedef double (*boolean_double)(int16_t);
typedef int32_t (*boolean_long)(int16_t);
typedef int16_t (*double_boolean)(double);
typedef int16_t (*long_boolean)(int32_t);
typedef int32_t (*long32)(int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t, int32_t);
struct wrong_thread {
    long2 call;
    double2 floating_call;
    boolean1 boolean_call;
    int32_t result;
    double floating_result;
    int16_t boolean_result;
    DWORD error, floating_error, boolean_error;
};
static DWORD WINAPI call_wrong_thread(void *pointer) {
    struct wrong_thread *probe = pointer;
    probe->result = probe->call(7,35);
    probe->error = GetLastError();
    probe->floating_result = probe->floating_call(1.25,2.5);
    probe->floating_error = GetLastError();
    probe->boolean_result = probe->boolean_call(-1);
    probe->boolean_error = GetLastError();
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
    double_long RoundLong = (double_long)GetProcAddress(dll,"RoundLong");
    double_long RoundThenAdd = (double_long)GetProcAddress(dll,"RoundThenAdd");
    boolean_double BooleanDouble = (boolean_double)GetProcAddress(dll,"BooleanDouble");
    boolean_long BooleanLong = (boolean_long)GetProcAddress(dll,"BooleanLong");
    double_boolean DoubleBoolean = (double_boolean)GetProcAddress(dll,"DoubleBoolean");
    long_boolean LongBoolean = (long_boolean)GetProcAddress(dll,"LongBoolean");
    long1 SumTo = (long1)GetProcAddress(dll,"SumTo");
    long1 Depth = (long1)GetProcAddress(dll,"Depth");
    long32 Weighted32 = (long32)GetProcAddress(dll,"Weighted32");
    if (!RoundLong || !RoundThenAdd || !BooleanDouble || !BooleanLong || !DoubleBoolean || !LongBoolean || !SumTo || !Depth || !Weighted32) return 10;
    long2 IntegerDivide = (long2)GetProcAddress(dll,"IntegerDivide");
    long2 IntegerMod = (long2)GetProcAddress(dll,"IntegerMod");
    boolean2 BooleanAnd = (boolean2)GetProcAddress(dll,"BooleanAnd");
    boolean2 BooleanOr = (boolean2)GetProcAddress(dll,"BooleanOr");
    boolean2 BooleanXor = (boolean2)GetProcAddress(dll,"BooleanXor");
    boolean2 BooleanEqv = (boolean2)GetProcAddress(dll,"BooleanEqv");
    boolean2 BooleanImp = (boolean2)GetProcAddress(dll,"BooleanImp");
    boolean1 BooleanNot = (boolean1)GetProcAddress(dll,"BooleanNot");
    long2 LongAnd = (long2)GetProcAddress(dll,"LongAnd");
    long2 LongOr = (long2)GetProcAddress(dll,"LongOr");
    long2 LongXor = (long2)GetProcAddress(dll,"LongXor");
    long2 LongEqv = (long2)GetProcAddress(dll,"LongEqv");
    long2 LongImp = (long2)GetProcAddress(dll,"LongImp");
    long1 LongNot = (long1)GetProcAddress(dll,"LongNot");
    if (!IntegerDivide || !IntegerMod || !BooleanAnd || !BooleanOr || !BooleanXor || !BooleanEqv || !BooleanImp || !BooleanNot || !LongAnd || !LongOr || !LongXor || !LongEqv || !LongImp || !LongNot) return 11;
    if (!add || !direct || !add_double || !greater || !echo || !mixed || !divide || !library_abs || !library_sqr) return 4;
    if ((FARPROC)add != GetProcAddress(dll,MAKEINTRESOURCEA(11))) return 5;
    printf("%d|%d|%.17g|%d|%d|%d|%d|%.17g|%.17g|%.17g|%.17g|%d|%.17g|%.17g|%.17g",
        add(-7,49),direct(5,32),add_double(1.25,2.5),
        greater(5,-2),greater(-2,5),echo(-1),echo(0),
        mixed(1.25,2,-1,3.5,4),mixed(1.25,2,0,3.5,4),
        divide(9,2),divide(9,0),add(INT32_MAX,1),library_abs(-12.5),library_sqr(25),library_sqr(-1));
    printf("|%d", RoundLong(-3.5));
    printf("|%d", RoundLong(-2.5));
    printf("|%d", RoundLong(-1.5));
    printf("|%d", RoundLong(-0.5));
    printf("|%d", RoundLong(0.5));
    printf("|%d", RoundLong(1.5));
    printf("|%d", RoundLong(2.5));
    printf("|%d", RoundLong(3.5));
    printf("|%d", RoundLong(2147483647.0));
    printf("|%d", RoundLong(2147483647.5));
    printf("|%d", RoundLong(-2147483648.5));
    printf("|%d", RoundLong(-2147483648.6));
    printf("|%d", RoundThenAdd(2.5));
    printf("|%d", RoundThenAdd(1e30));
    printf("|%.17g", BooleanDouble(-1));
    printf("|%.17g", BooleanDouble(0));
    printf("|%d", BooleanLong(-1));
    printf("|%d", BooleanLong(0));
    printf("|%d", DoubleBoolean(-0.5));
    printf("|%d", DoubleBoolean(0.0));
    printf("|%d", DoubleBoolean(1e-100));
    printf("|%d", LongBoolean(-2147483647));
    printf("|%d", LongBoolean(0));
    printf("|%d", LongBoolean(1));
    printf("|%d", SumTo(10));
    printf("|%d", SumTo(65535));
    printf("|%d", SumTo(65536));
    printf("|%d", Depth(32));
    printf("|%d", Depth(127));
    printf("|%d", Depth(128));
    printf("|%d", Depth(10000));
    printf("|%d", Depth(32));
    printf("|%d", Weighted32(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32));
    printf("|%d", Weighted32(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -7));
    printf("|%d", IntegerDivide(-7, 3));
    printf("|%d", IntegerDivide(7, -3));
    printf("|%d", IntegerDivide(INT32_MIN, -1));
    printf("|%d", IntegerDivide(9, 0));
    printf("|%d", IntegerDivide(INT32_MIN, 1));
    printf("|%d", IntegerMod(-7, 3));
    printf("|%d", IntegerMod(7, -3));
    printf("|%d", IntegerMod(INT32_MIN, -1));
    printf("|%d", IntegerMod(9, 0));
    printf("|%d", IntegerMod(INT32_MIN, 1));
    printf("|%d", BooleanAnd(0, 0));
    printf("|%d", BooleanAnd(0, -1));
    printf("|%d", BooleanAnd(-1, 0));
    printf("|%d", BooleanAnd(-1, -1));
    printf("|%d", BooleanOr(0, 0));
    printf("|%d", BooleanOr(0, -1));
    printf("|%d", BooleanOr(-1, 0));
    printf("|%d", BooleanOr(-1, -1));
    printf("|%d", BooleanXor(0, 0));
    printf("|%d", BooleanXor(0, -1));
    printf("|%d", BooleanXor(-1, 0));
    printf("|%d", BooleanXor(-1, -1));
    printf("|%d", BooleanEqv(0, 0));
    printf("|%d", BooleanEqv(0, -1));
    printf("|%d", BooleanEqv(-1, 0));
    printf("|%d", BooleanEqv(-1, -1));
    printf("|%d", BooleanImp(0, 0));
    printf("|%d", BooleanImp(0, -1));
    printf("|%d", BooleanImp(-1, 0));
    printf("|%d", BooleanImp(-1, -1));
    printf("|%d", BooleanNot(0));
    printf("|%d", BooleanNot(-1));
    printf("|%d", LongAnd(-123, 456));
    printf("|%d", LongAnd(INT32_MIN, 65535));
    printf("|%d", LongOr(-123, 456));
    printf("|%d", LongOr(INT32_MIN, 65535));
    printf("|%d", LongXor(-123, 456));
    printf("|%d", LongXor(INT32_MIN, 65535));
    printf("|%d", LongEqv(-123, 456));
    printf("|%d", LongEqv(INT32_MIN, 65535));
    printf("|%d", LongImp(-123, 456));
    printf("|%d", LongImp(INT32_MIN, 65535));
    printf("|%d", LongNot(INT32_MIN));
    printf("|%d", LongNot(0));
    printf("|%d", LongNot(2147483647));
    printf("\n");
    uint32_t seed = 0x3917a0d3u;
    for (int iteration = 0; iteration < 4096; ++iteration) {
        seed ^= seed << 13; seed ^= seed >> 17; seed ^= seed << 5;
        int32_t a = (int32_t)seed;
        seed ^= seed << 13; seed ^= seed >> 17; seed ^= seed << 5;
        int32_t b = (int32_t)seed;
        int64_t wide = (int64_t)a + b;
        int32_t expected_add = wide < INT32_MIN || wide > INT32_MAX ? 0 : (int32_t)wide;
        int64_t quotient = b == 0 ? 0 : (int64_t)a / b;
        int32_t expected_divide = quotient < INT32_MIN || quotient > INT32_MAX ? 0 : (int32_t)quotient;
        int32_t expected_mod = b == 0 ? 0 : (int32_t)((int64_t)a % b);
        int32_t rounded = a / 2;
        if (a % 2 != 0 && rounded % 2 != 0) rounded += a > 0 ? 1 : -1;
        if (add(a,b) != expected_add || IntegerDivide(a,b) != expected_divide || IntegerMod(a,b) != expected_mod ||
            LongAnd(a,b) != (a & b) || LongOr(a,b) != (a | b) || LongXor(a,b) != (a ^ b) ||
            LongEqv(a,b) != ~(a ^ b) || LongImp(a,b) != (~a | b) || LongNot(a) != ~a ||
            LongBoolean(a) != (a == 0 ? 0 : -1) || RoundLong((double)a / 2.0) != rounded ||
            RoundThenAdd((double)a / 2.0) != rounded + 1) return 12;
        int16_t first = a < 0 ? -1 : 0;
        int16_t second = b < 0 ? -1 : 0;
        if (BooleanAnd(first,second) != (first & second) || BooleanOr(first,second) != (first | second) ||
            BooleanXor(first,second) != (first ^ second) || BooleanEqv(first,second) != ~(first ^ second) ||
            BooleanImp(first,second) != (~first | second) || BooleanNot(first) != ~first) return 13;
    }
    if (add_double(NAN,1) != 0 || library_abs(INFINITY) != 0) return 6;
    struct wrong_thread probe = {0};
    probe.call = add;
    probe.floating_call = add_double;
    probe.boolean_call = echo;
    HANDLE thread = CreateThread(NULL,0,call_wrong_thread,&probe,0,NULL);
    if (!thread) return 7;
    if (WaitForSingleObject(thread,5000) != WAIT_OBJECT_0) return 8;
    CloseHandle(thread);
    if (probe.result != 0 || probe.error != 1444 || probe.floating_result != 0 || probe.floating_error != 1444 ||
        probe.boolean_result != 0 || probe.boolean_error != 1444) return 9;
    if (add(7,35) != 42) return 14;
    FreeLibrary(dll);
    // A fresh mapping may claim a different owner. Once that owner exits,
    // later threads must still be refused; there is no per-thread context.
    dll = LoadLibraryW(argv[1]);
    if (!dll) return 15;
    probe.call = (long2)GetProcAddress(dll,"AddLong");
    probe.floating_call = (double2)GetProcAddress(dll,"AddDouble");
    probe.boolean_call = (boolean1)GetProcAddress(dll,"EchoBoolean");
    if (!probe.call || !probe.floating_call || !probe.boolean_call) return 16;
    thread = CreateThread(NULL,0,call_wrong_thread,&probe,0,NULL);
    if (!thread || WaitForSingleObject(thread,5000) != WAIT_OBJECT_0) return 17;
    CloseHandle(thread);
    if (probe.result != 42 || probe.floating_result != 3.75 || probe.boolean_result != -1) return 18;
    SetLastError(0);
    if (probe.call(7,35) != 0 || GetLastError() != 1444) return 19;
    for (int attempt = 0; attempt < 16; ++attempt) {
        thread = CreateThread(NULL,0,call_wrong_thread,&probe,0,NULL);
        if (!thread || WaitForSingleObject(thread,5000) != WAIT_OBJECT_0) return 20;
        CloseHandle(thread);
        if (probe.result != 0 || probe.error != 1444 || probe.floating_result != 0 || probe.floating_error != 1444 ||
            probe.boolean_result != 0 || probe.boolean_error != 1444) return 21;
    }
    FreeLibrary(dll);
    dll = LoadLibraryW(argv[1]);
    if (!dll) return 22;
    add = (long2)GetProcAddress(dll,"AddLong");
    if (!add || add(7,35) != 42) return 23;
    FreeLibrary(dll);
    return 0;
}
