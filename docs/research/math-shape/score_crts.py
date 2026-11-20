# dc-q6brh.112 shape check: scores CRT DLLs on this Windows machine against Excel VBA's answers
# (MathShape.excel.txt, from MathShape.bas, 2026-11-20). msvcr100/msvcr110 match all 9567 cases;
# VBE7.DLL imports MSVCR100.dll. Tan is Sin/Cos. Run: py -3 -I score_crts.py
import json, struct, ctypes, collections, os
import pathlib
text = (pathlib.Path(__file__).parent / 'MathShape.excel.txt').read_text()
def d(h): return struct.unpack('>d', bytes.fromhex(h))[0]
def hx(v): return struct.pack('>d', v).hex().upper()
libs = {}
for n in ['msvcr100', 'msvcr110', 'msvcr120', r'C:\Program Files\Microsoft Office\root\Office16\UCRTBASE.DLL']:
    try: libs[os.path.basename(n)] = ctypes.CDLL(n)
    except OSError as e: print('skip', n, e)
fn = {'sin': 'sin', 'cos': 'cos', 'atn': 'atan', 'exp': 'exp', 'log': 'log'}
for l in libs.values():
    for f in fn.values():
        getattr(l, f).restype = ctypes.c_double; getattr(l, f).argtypes = [ctypes.c_double]
lines = [l.split() for l in text.splitlines() if not l.split()[2].startswith('ERR')]
for name, l in libs.items():
    for mode in ([1, 0] if hasattr(l, '_set_FMA3_enable') else [None]):
        if mode is not None: l._set_FMA3_enable(mode)
        c = collections.Counter()
        for f, xh, yh in lines:
            x = d(xh); c[f + 'n'] += 1
            y = l.sin(x) / l.cos(x) if f == 'tan' else getattr(l, fn[f])(x)
            c[f] += hx(y) == yh
        print(name, 'FMA3', mode, {f: f"{c[f]}/{c[f + 'n']}" for f in ['sin', 'cos', 'tan', 'atn', 'exp', 'log']})
