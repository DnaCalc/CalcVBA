# Dumps the VBA standard library as Excel's VBA sees it: the type library inside VBE7.DLL
# (loaded with REGKIND_NONE, so nothing is registered). Writes vbe7-library.tsv:
# typeinfo, kind, member, member kind, flags, return type, parameters (name:type, ? = optional).
#
#   pwsh -File dump-vbe7-typelib.ps1 [-Dll <path to VBE7.DLL>]
param(
    [string]$Dll = 'C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL'
)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

public static class TlbDump {
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void LoadTypeLibEx(string file, int regkind, out ITypeLib tlib);

    static string Vt(ITypeInfo ti, TYPEDESC td) {
        VarEnum vt = (VarEnum)td.vt;
        switch (vt) {
            case VarEnum.VT_PTR: return Vt(ti, (TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(TYPEDESC))) + "*";
            case VarEnum.VT_SAFEARRAY: return Vt(ti, (TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(TYPEDESC))) + "()";
            case VarEnum.VT_USERDEFINED: {
                ITypeInfo r; ti.GetRefTypeInfo(unchecked((int)td.lpValue.ToInt64()), out r);
                string n, d, h; int c; r.GetDocumentation(-1, out n, out d, out c, out h); return n;
            }
            default: return vt.ToString().Replace("VT_", "");
        }
    }

    public static string Dump(string path) {
        ITypeLib tl; LoadTypeLibEx(path, 2, out tl);
        var sb = new StringBuilder();
        sb.Append("typeinfo\tkind\tflags\tmember\tmember_kind\tmember_flags\treturns\tparams\tvalue\n");
        int count = tl.GetTypeInfoCount();
        for (int i = 0; i < count; i++) {
            ITypeInfo ti; tl.GetTypeInfo(i, out ti);
            string tname, tdoc, thelp; int tctx; ti.GetDocumentation(-1, out tname, out tdoc, out tctx, out thelp);
            IntPtr pa; ti.GetTypeAttr(out pa);
            TYPEATTR ta = (TYPEATTR)Marshal.PtrToStructure(pa, typeof(TYPEATTR));
            string tkind = ta.typekind.ToString().Replace("TKIND_", "");
            string tflags = ((int)ta.wTypeFlags).ToString("X");
            if (ta.cFuncs == 0 && ta.cVars == 0) sb.Append(tname + "\t" + tkind + "\t" + tflags + "\t\t\t\t\t\t\n");
            for (int f = 0; f < ta.cFuncs; f++) {
                IntPtr pf; ti.GetFuncDesc(f, out pf);
                FUNCDESC fd = (FUNCDESC)Marshal.PtrToStructure(pf, typeof(FUNCDESC));
                string[] names = new string[fd.cParams + 1]; int got;
                ti.GetNames(fd.memid, names, names.Length, out got);
                var ps = new List<string>();
                int size = Marshal.SizeOf(typeof(ELEMDESC));
                for (int p = 0; p < fd.cParams; p++) {
                    ELEMDESC ed = (ELEMDESC)Marshal.PtrToStructure(new IntPtr(fd.lprgelemdescParam.ToInt64() + p * size), typeof(ELEMDESC));
                    PARAMFLAG pfl = ed.desc.paramdesc.wParamFlags;
                    bool opt = (pfl & PARAMFLAG.PARAMFLAG_FOPT) != 0;
                    string pn = (p + 1 < got) ? names[p + 1] : ("p" + p);
                    ps.Add((opt ? "?" : "") + pn + ":" + Vt(ti, ed.tdesc));
                }
                if (fd.cParamsOpt == -1) ps.Add("ParamArray");
                string kind = fd.invkind.ToString().Replace("INVOKE_", "");
                sb.Append(tname + "\t" + tkind + "\t" + tflags + "\t" + names[0] + "\t" + kind + "\t" +
                    ((int)fd.wFuncFlags).ToString("X") + "\t" + Vt(ti, fd.elemdescFunc.tdesc) + "\t" + string.Join(", ", ps) + "\t\n");
                ti.ReleaseFuncDesc(pf);
            }
            for (int v = 0; v < ta.cVars; v++) {
                IntPtr pv; ti.GetVarDesc(v, out pv);
                VARDESC vd = (VARDESC)Marshal.PtrToStructure(pv, typeof(VARDESC));
                string[] names = new string[1]; int got; ti.GetNames(vd.memid, names, 1, out got);
                string val = "";
                if (vd.varkind == VARKIND.VAR_CONST) {
                    object o = Marshal.GetObjectForNativeVariant(vd.desc.lpvarValue);
                    val = o == null ? "" : (o.GetType().Name + ":" + Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture));
                }
                sb.Append(tname + "\t" + tkind + "\t" + tflags + "\t" + names[0] + "\t" + vd.varkind.ToString().Replace("VAR_", "") + "\t" +
                    ((int)vd.wVarFlags).ToString("X") + "\t" + Vt(ti, vd.elemdescVar.tdesc) + "\t\t" + val + "\n");
                ti.ReleaseVarDesc(pv);
            }
            ti.ReleaseTypeAttr(pa);
        }
        return sb.ToString();
    }
}
'@
$out = Join-Path $PSScriptRoot 'vbe7-library.tsv'
[IO.File]::WriteAllText($out, [TlbDump]::Dump($Dll).Replace("`r`n", "`n"))
"wrote $out"
