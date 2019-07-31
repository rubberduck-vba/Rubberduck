using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    /// <summary>
    /// A compatible version of ComTypes.ITypeInfo, using IntPtr for all out params
    /// see https://msdn.microsoft.com/en-gb/library/windows/desktop/ms221696(v=vs.85).aspx
    /// </summary>
    /// <remarks>
    /// We use [PreserveSig] (<see cref="PreserveSigAttribute"/>) so that we can handle HRESULTs directly
    /// </remarks>
    [ComImport(), Guid("00020401-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITypeInfoInternal
    {
        [PreserveSig] /*HRESULT*/ int GetTypeAttr(/*out TYPEATTR*/ IntPtr ppTypeAttr);
        [PreserveSig] /*HRESULT*/ int GetTypeComp(/*out ITypeComp*/ IntPtr ppTComp);
        [PreserveSig] /*HRESULT*/ int GetFuncDesc(int index, /*out FUNCDESC*/ IntPtr ppFuncDesc);
        [PreserveSig] /*HRESULT*/ int GetVarDesc(int index, /*out VARDESC*/ IntPtr ppVarDesc);
        [PreserveSig] /*HRESULT*/ int GetNames(int memid, /*string[]*/ IntPtr rgBstrNames, int cMaxNames, /*out int*/ IntPtr pcNames);
        [PreserveSig] /*HRESULT*/ int GetRefTypeOfImplType(int index, /*out HREFTYPE*/ IntPtr href);
        [PreserveSig] /*HRESULT*/ int GetImplTypeFlags(int index, /*out ComTypes.IMPLTYPEFLAGS*/ IntPtr pImplTypeFlags);
        [PreserveSig] /*HRESULT*/ int GetIDsOfNames(/*string[]*/ IntPtr rgszNames, int cNames, /*out MEMBERID*/ IntPtr pMemId);
        [PreserveSig] /*HRESULT*/ int Invoke(/*object*/ IntPtr pvInstance, int memid, short wFlags,/*ref ComTypes.DISPPARAMS*/ IntPtr pDispParams, /*out VARIANT*/ IntPtr pVarResult, /*out EXCEPINFO*/ IntPtr pExcepInfo, /*out int*/ IntPtr puArgErr);
        [PreserveSig] /*HRESULT*/ int GetDocumentation(int index, /*out string*/ IntPtr strName, /*out string*/ IntPtr strDocString, /*out int*/ IntPtr dwHelpContext, /*out string*/ IntPtr strHelpFile);
        [PreserveSig] /*HRESULT*/ int GetDllEntry(int memid, System.Runtime.InteropServices.ComTypes.INVOKEKIND invKind, /*out string*/ IntPtr pBstrDllName, /*out string*/IntPtr pBstrName, /*out short*/ IntPtr pwOrdinal);
        [PreserveSig] /*HRESULT*/ int GetRefTypeInfo(int hRef, /*out ITypeInfo*/ IntPtr ppTI);
        [PreserveSig] /*HRESULT*/ int AddressOfMember(int memid, System.Runtime.InteropServices.ComTypes.INVOKEKIND invKind, /*out IntPtr*/ IntPtr ppv);
        [PreserveSig] /*HRESULT*/ int CreateInstance(/*object*/ IntPtr pUnkOuter, ref Guid riid, /*out IntPtr*/ IntPtr ppvObj);
        [PreserveSig] /*HRESULT*/ int GetMops(int memid, /*out string*/ IntPtr pBstrMops);
        [PreserveSig] /*HRESULT*/ int GetContainingTypeLib(/*out ITypeLib*/ IntPtr ppTLB, /*out int*/ IntPtr pIndex);
        [PreserveSig] void ReleaseTypeAttr(/*TYPEATTR*/ IntPtr pTypeAttr);
        [PreserveSig] void ReleaseFuncDesc(/*FUNCDESC*/ IntPtr pFuncDesc);
        [PreserveSig] void ReleaseVarDesc(/*VARDESC*/ IntPtr pVarDesc);
    }
}