using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoWrapper : ITypeInfo, IDisposable
    {
        ITypeLib Container { get; }
        int ContainerIndex { get; }
        bool HasModuleScopeCompilationErrors { get; }
        bool HasVBEExtensions { get; }
        System.Runtime.InteropServices.ComTypes.TYPEATTR CachedAttributes { get; }
        bool HasSimulatedContainer { get; }
        bool IsUserFormBaseClass { get; }
        string Name { get; }
        string DocString { get; }
        int HelpContext { get; }
        string HelpFile { get; }
        string ProgID { get; }
        Guid GUID { get; }
        TYPEKIND_VBE TypeKind { get; }
        bool HasPredeclaredId { get; }
        System.Runtime.InteropServices.ComTypes.TYPEFLAGS Flags { get; }
        string ContainerName { get; }
        ITypeInfoVBEExtensions VBEExtensions { get; }
        //void Dispose();
        int GetSafeRefTypeInfo(int hRef, out ITypeInfoWrapper outTI);
        IntPtr GetCOMReferencePtr();
        int GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex);
        int GetTypeAttr(IntPtr ppTypeAttr);
        int GetTypeComp(IntPtr ppTComp);
        int GetFuncDesc(int index, IntPtr ppFuncDesc);
        int GetVarDesc(int index, IntPtr ppVarDesc);
        int GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames);
        int GetRefTypeOfImplType(int index, IntPtr href);
        int GetImplTypeFlags(int index, IntPtr pImplTypeFlags);
        int GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId);
        int Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);
        int GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile);
        new int GetDllEntry(int memid, System.Runtime.InteropServices.ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        int GetRefTypeInfo(int hRef, IntPtr ppTI);
        int AddressOfMember(int memid, System.Runtime.InteropServices.ComTypes.INVOKEKIND invKind, IntPtr ppv);
        int CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj);
        int GetMops(int memid, IntPtr pBstrMops);
        new void ReleaseTypeAttr(IntPtr pTypeAttr);
        new void ReleaseFuncDesc(IntPtr pFuncDesc);
        new void ReleaseVarDesc(IntPtr pVarDesc);
        ITypeInfoFunctionCollection Funcs { get; }
        ITypeInfoVariablesCollection Vars { get; }
        ITypeInfoImplementedInterfacesCollection ImplementedInterfaces { get; }
    }
}