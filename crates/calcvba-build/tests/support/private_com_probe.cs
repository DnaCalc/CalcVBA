using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class CalcVbaPrivateComProbe {
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] public delegate int ServerCall();
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate int GetFactory(ref Guid clsid, ref Guid iid, out IntPtr value);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate int Create(IntPtr self, IntPtr outer, ref Guid iid, out IntPtr value);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate int Lock(IntPtr self, int value);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate uint Release(IntPtr self);
    [DllImport("ole32.dll")] private static extern int CoInitializeEx(IntPtr reserved,uint flags);
    [DllImport("ole32.dll")] private static extern void CoUninitialize();
    [DllImport("oleaut32.dll",CharSet=CharSet.Unicode)] private static extern int LoadTypeLibEx(string path,int mode,out ITypeLib value);
    private static IntPtr image;
    public static void Load(string path) { image=NativeLibrary.Load(path); }
    public static int Call(string name) { return Marshal.GetDelegateForFunctionPointer<ServerCall>(NativeLibrary.GetExport(image,name))(); }
    private static T Slot<T>(IntPtr value,int slot) where T:Delegate {
        return Marshal.GetDelegateForFunctionPointer<T>(Marshal.ReadIntPtr(Marshal.ReadIntPtr(value),slot*IntPtr.Size));
    }
    private static void Check(bool condition,string message) {if(!condition) throw new Exception(message);}
    public static void Factories(string[] ids,bool[] creatable) {
        int initialized=CoInitializeEx(IntPtr.Zero,2); Check(initialized>=0,"STA initialization failed");
        try {
            GetFactory get=Marshal.GetDelegateForFunctionPointer<GetFactory>(NativeLibrary.GetExport(image,"DllGetClassObject"));
            Guid factoryId=new Guid("00000001-0000-0000-C000-000000000046");
            Guid dispatchId=new Guid("00020400-0000-0000-C000-000000000046");
            for(int i=0;i<ids.Length;i++) {
                Guid clsid=new Guid(ids[i]); IntPtr factory; int hr=get(ref clsid,ref factoryId,out factory);
                if(!creatable[i]) {Check(hr==unchecked((int)0x80040111)&&factory==IntPtr.Zero,"noncreatable factory exposed");continue;}
                Check(hr==0&&factory!=IntPtr.Zero,"class factory failed");
                Check(Call("DllCanUnloadNow")==1,"live factory must retain DLL");
                Lock locker=Slot<Lock>(factory,4); Check(locker(factory,1)==0,"LockServer failed");
                Slot<Release>(factory,2)(factory);
                Check(Call("DllCanUnloadNow")==1,"LockServer did not retain DLL");
                Check(get(ref clsid,ref factoryId,out factory)==0,"locked factory reacquisition failed");
                Create create=Slot<Create>(factory,3); IntPtr first=IntPtr.Zero,second=IntPtr.Zero;
                try {
                    Check(create(factory,new IntPtr(1),ref dispatchId,out first)==unchecked((int)0x80040110)&&first==IntPtr.Zero,"aggregation accepted");
                    Guid unknown=new Guid("69E9525E-28AE-4014-A81C-9925ACD034D3");
                    Check(create(factory,IntPtr.Zero,ref unknown,out first)==unchecked((int)0x80004002)&&first==IntPtr.Zero,"unknown interface accepted");
                    Check(create(factory,IntPtr.Zero,ref dispatchId,out first)==0,"first instance failed");
                    Check(create(factory,IntPtr.Zero,ref dispatchId,out second)==0&&first!=second,"instances share identity");
                } finally {
                    if(first!=IntPtr.Zero) Slot<Release>(first,2)(first);
                    if(second!=IntPtr.Zero) Slot<Release>(second,2)(second);
                    Check(locker(factory,0)==0,"UnlockServer failed");
                    Slot<Release>(factory,2)(factory);
                }
                Check(Call("DllCanUnloadNow")==0,"COM references leaked");
            }
        } finally { CoUninitialize(); }
    }
    public static Dictionary<string,int> TypeFlags(string path) {
        ITypeLib library; int hr=LoadTypeLibEx(path,2,out library); Check(hr==0,"LoadTypeLibEx failed");
        var result=new Dictionary<string,int>();
        try {
            for(int i=0;i<library.GetTypeInfoCount();i++) {
                TYPEKIND kind; library.GetTypeInfoType(i,out kind); if(kind!=TYPEKIND.TKIND_COCLASS) continue;
                ITypeInfo info; library.GetTypeInfo(i,out info); IntPtr attr=IntPtr.Zero;
                try {
                    string name,doc,file;int context;info.GetDocumentation(-1,out name,out doc,out context,out file);
                    info.GetTypeAttr(out attr); result[name]=(int)Marshal.PtrToStructure<TYPEATTR>(attr).wTypeFlags;
                } finally {if(attr!=IntPtr.Zero) info.ReleaseTypeAttr(attr);Marshal.FinalReleaseComObject(info);}
            }
        } finally {Marshal.FinalReleaseComObject(library);}
        return result;
    }
}
