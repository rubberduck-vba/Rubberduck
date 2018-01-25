using System;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibsAbstract
{
    [ComImport(), Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDispatch
    {
    }

    // A compatible version of ITypeInfo, where COM objects are outputted as IntPtrs instead of objects
    [ComImport(), Guid("00020401-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITypeInfo_Ptrs
    {
        void GetTypeAttr(out IntPtr ppTypeAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetFuncDesc(int index, out IntPtr ppFuncDesc);
        void GetVarDesc(int index, out IntPtr ppVarDesc);
        void GetNames(int memid, [Out] out string rgBstrNames, int cMaxNames, out int pcNames);
        void GetRefTypeOfImplType(int index, out int href);
        void GetImplTypeFlags(int index, out ComTypes.IMPLTYPEFLAGS pImplTypeFlags);
        void GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId);
        void Invoke(object pvInstance, int memid, short wFlags, ref ComTypes.DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr);
        void GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile);
        void GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        void GetRefTypeInfo(int hRef, out IntPtr ppTI);
        void AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, out IntPtr ppv);
        void CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj);
        void GetMops(int memid, out string pBstrMops);
        void GetContainingTypeLib(out IntPtr ppTLB, out int pIndex);
        void ReleaseTypeAttr(IntPtr pTypeAttr);
        void ReleaseFuncDesc(IntPtr pFuncDesc);
        void ReleaseVarDesc(IntPtr pVarDesc);
    }

    [ComImport(), Guid("DDD557E1-D96F-11CD-9570-00AA0051E5D4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IVBEComponent
    {
        void Placeholder1();
        void Placeholder2();
        void Placeholder3();
        void Placeholder4();
        void Placeholder5();
        void Placeholder6();
        void Placeholder7();
        void Placeholder8();
        void Placeholder9();
        void Placeholder10();
        void Placeholder11();
        void Placeholder12();
        void CompileComponent();
        void Placeholder14();
        void Placeholder15();
        void Placeholder16();
        void Placeholder17();
        void Placeholder18();
        void Placeholder19();
        void Placeholder20();
        void Placeholder21();
        void Placeholder22();
        void Placeholder23();
        void Placeholder24();
        void Placeholder25();
        void Placeholder26();
        void Placeholder27();
        void Placeholder28();
        void Placeholder29();
        void Placeholder30();
        void Placeholder31();
        void Placeholder32();
        void Placeholder33();
        void GetSomeRelatedTypeInfoPtrs(out IntPtr A, out IntPtr B);        // returns 2 TypeInfos, seemingly related to this ITypeInfo, but slightly different.
    }

    // An extended version of ITypeInfo, hosted by the VBE that includes a particularly helpful member, GetStdModInstance
    [ComImport(), Guid("CACC1E82-622B-11D2-AA78-00C04F9901D2")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IVBETypeInfo
    {
        void GetTypeAttr(out IntPtr ppTypeAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetFuncDesc(int index, out IntPtr ppFuncDesc);
        void GetVarDesc(int index, out IntPtr ppVarDesc);
        void GetNames(int memid, [Out] out string rgBstrNames, int cMaxNames, out int pcNames);
        void GetRefTypeOfImplType(int index, out int href);
        void GetImplTypeFlags(int index, out ComTypes.IMPLTYPEFLAGS pImplTypeFlags);
        void GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId);
        void Invoke(object pvInstance, int memid, short wFlags, ref ComTypes.DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr);
        void GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile);
        void GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        void GetRefTypeInfo(int hRef, out IntPtr ppTI);
        void AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, out IntPtr ppv);
        void CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj);
        void GetMops(int memid, out string pBstrMops);
        void GetContainingTypeLib(out IntPtr ppTLB, out int pIndex);
        void ReleaseTypeAttr(IntPtr pTypeAttr);
        void ReleaseFuncDesc(IntPtr pFuncDesc);
        void ReleaseVarDesc(IntPtr pVarDesc);

        void Placeholder1();
        IDispatch GetStdModInstance();            // a handy extra vtable entry we can use to invoke members in standard modules.
    }

    // A compatible version of ITypeLib, where COM objects are outputted as IntPtrs instead of objects
    [ComImport(), Guid("00020402-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITypeLib_Ptrs
    {
        int GetTypeInfoCount();
        void GetTypeInfo(int index, out IntPtr ppTI);
        void GetTypeInfoType(int index, out ComTypes.TYPEKIND pTKind);
        void GetTypeInfoOfGuid(ref Guid guid, out IntPtr ppTInfo);
        void GetLibAttr(out IntPtr ppTLibAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile);
        bool IsName(string szNameBuf, int lHashVal);
        void FindName(string szNameBuf, int lHashVal, IntPtr[] ppTInfo, int[] rgMemId, ref short pcFound);
        void ReleaseTLibAttr(IntPtr pTLibAttr);
    }

    // An internal representation of the VBE References collection object, as returned from the VBE.ActiveVBProject.References, or similar
    // These offsets are known to be valid across 32-bit and 64-bit versions of VBA and VB6, right back from when VBA6 was first released.
    [StructLayout(LayoutKind.Sequential)]
    struct VBEReferencesObj
    {
        IntPtr vTable1;     // _References vtable
        IntPtr vTable2;
        IntPtr vTable3;
        IntPtr Object1;
        IntPtr Object2;
        public IntPtr TypeLib;
        IntPtr Placeholder1;
        IntPtr Placeholder2;
        IntPtr RefCount;
    }

    // A ITypeLib object hosted by the VBE, also providing Prev/Next pointers for a double linked list of all loaded project ITypeLibs
    [StructLayout(LayoutKind.Sequential)]
    struct VBETypeLibObj
    {
        IntPtr vTable1;     // ITypeLib vtable
        IntPtr vTable2;
        IntPtr vTable3;
        public IntPtr Prev;
        public IntPtr Next;
    }

    // IVBEProject, obtainable from a VBE hosted ITypeLib in order to access a few extra features...
    [ComImport(), Guid("DDD557E0-D96F-11CD-9570-00AA0051E5D4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IVBEProject
    {
        string GetProjectName();                 // same as calling ITypeLib::GetDocumentation(-1)                   
        void SetProjectName(string value);       // same as IVBEProject2::set_ProjectName()
        int GetVbeLCID();
        void Placeholder3();                      // calls IVBEProject2::Placeholder8
        void Placeholder4();                      
        void Placeholder5();                    
        void Placeholder6();
        void Placeholder7();
        string GetConditionalCompilationArgs();
        void SetConditionalCompilationArgs(string args);
        void Placeholder8();
        void Placeholder9();
        void Placeholder10();
        void Placeholder11();
        void Placeholder12();
        void Placeholder13();
        void Placeholder14();
        void Placeholder15();
        void Placeholder16();
        void Placeholder17();
        string GetReferenceString(int ReferenceIndex); // the raw reference string
        void CompileProject();                            // throws COM exception 0x800A9C64 if error occurred during compile.
    }

    // IVBEProject2, vtable position just before the IVBEProject, not queryable, so needs aggregation
    [ComImport(), Guid("FFFFFFFF-0000-0000-C000-000000000046")]  // 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IVBEProject2
    {
        void Placeholder1();                    // returns E_NOTIMPL
        void SetProjectName(string value);
        void SetProjectVersion(ushort wMajorVerNum, ushort wMinorVerNum);
        void SetProjectGUID(ref Guid value);
        void SetProjectDescription(string value);
        void SetProjectHelpFileName(string value);
        void SetProjectHelpContext(int value);
    }

    public enum TYPEKIND_VBE
    {
        TKIND_ENUM = 0,
        TKIND_RECORD = 1,
        TKIND_MODULE = 2,
        TKIND_INTERFACE = 3,
        TKIND_DISPATCH = 4,
        TKIND_COCLASS = 5,
        TKIND_ALIAS = 6,
        TKIND_UNION = 7,

        TKIND_VBACLASS = 8,                 // extended by VBA, this is used for the outermost interface
    }

    public enum TypeLibConsts : int
    {
        MEMBERID_NIL = -1,
    }

    public enum VBECompilerConsts : int
    {
        E_VBA_COMPILEERROR = unchecked((int)0x800A9C64)
    }
}
