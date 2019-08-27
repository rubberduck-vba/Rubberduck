using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    /// <summary>
    /// A version of IDispatch that allows us to call its members explicitly
    /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221608(v=vs.85).aspx
    /// </summary>
    [ComImport(), Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        [PreserveSig] int GetTypeInfoCount([Out] out uint pctinfo);
        [PreserveSig] int GetTypeInfo([In] uint iTInfo, [In] uint lcid, [Out] out System.Runtime.InteropServices.ComTypes.ITypeInfo pTypeInfo);
        [PreserveSig] int GetIDsOfNames([In] ref Guid riid, [In] string[] rgszNames, [In] uint cNames, [In] uint lcid, [Out] out int[] rgDispId);

        [PreserveSig]
        int Invoke([In] int dispIdMember,
            [In] ref Guid riid,
            [In] uint lcid,
            [In] uint dwFlags,
            [In, Out] ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
            [Out] out object pVarResult,
            [In, Out] ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
            [Out] out uint pArgErr);
    }
}