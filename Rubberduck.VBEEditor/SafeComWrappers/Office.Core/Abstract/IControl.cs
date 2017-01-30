using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface IControl : ISafeComWrapper, IEquatable<IControl>
    {
        string Name { get; set; }
    }

    public static class ControlExtensions
    {
        public static string TypeName(this IControl control)
        {
            try
            {
                var dispatch = control.Target as IDispatch;
                if (dispatch == null)
                {
                    return "Control";
                }
                ITypeInfo info;               
                dispatch.GetTypeInfo(0, 0, out info);
                string name;
                string docs;
                int context;
                string help;
                info.GetDocumentation(-1, out name, out docs, out context, out help);
                return name;
            }
            catch
            {
                return "Control";
            }
        }

        [ComImport]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IDispatch
        {
            [PreserveSig]
            int GetTypeInfoCount(out int Count);

            [PreserveSig]
            int GetTypeInfo(
                [MarshalAs(UnmanagedType.U4)] int iTInfo,
                [MarshalAs(UnmanagedType.U4)] int lcid,
                out ITypeInfo typeInfo);

            [PreserveSig]
            int GetIDsOfNames(
                ref Guid riid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)]
			string[] rgsNames,
                int cNames,
                int lcid,
                [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

            [PreserveSig]
            int Invoke(
                int dispIdMember,
                ref Guid riid,
                uint lcid,
                ushort wFlags,
                ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
                out object pVarResult,
                ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
                IntPtr[] pArgErr);
        }        
    }
}