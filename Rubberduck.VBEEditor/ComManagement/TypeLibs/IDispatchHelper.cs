using System;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Used by methods in the ITypeInfo and ITypeLib interfaces.  Usually used to get the root type or library name.
    /// </summary>
    public enum KnownDispatchMemberIDs : int
    {
        MEMBERID_NIL = -1,
    }

    /// <summary>
    /// Simplified equivalent of VARIANT structure often used in COM
    /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221627(v=vs.85).aspx
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct VARIANT
    {
        short _vt;
        short _reserved1;
        short _reserved2;
        short _reserved3;
        IntPtr _data1;
        IntPtr _data2;
    }

    /// <summary>
    /// A version of IDispatch that allows us to call its members explicitly
    /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221608(v=vs.85).aspx
    /// </summary>
    [ComImport(), Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        [PreserveSig] int GetTypeInfoCount([Out] out uint pctinfo);
        [PreserveSig] int GetTypeInfo([In] uint iTInfo, [In] uint lcid, [Out] out ComTypes.ITypeInfo pTypeInfo);
        [PreserveSig] int GetIDsOfNames([In] ref Guid riid, [In] string[] rgszNames, [In] uint cNames, [In] uint lcid, [Out] out int[] rgDispId);

        [PreserveSig]
        int Invoke([In] int dispIdMember,
            [In] ref Guid riid,
            [In] uint lcid,
            [In] uint dwFlags,
            [In, Out] ref ComTypes.DISPPARAMS pDispParams,
            [Out] out Object pVarResult,
            [In, Out] ref ComTypes.EXCEPINFO pExcepInfo,
            [Out] out uint pArgErr);
    }

    public static class IDispatchHelper
    {
        private static Guid _guid_null = new Guid();

        /// <summary>
        /// IDispatch::Invoke flags
        /// see https://msdn.microsoft.com/en-gb/library/windows/desktop/ms221479(v=vs.85).aspx
        /// </summary>
        public enum InvokeKind : int
        {
            DISPATCH_METHOD = 1,
            DISPATCH_PROPERTYGET = 2,
            DISPATCH_PROPERTYPUT = 4,
            DISPATCH_PROPERTYPUTREF = 8,
        }

        /// <summary>
        /// Convert input args into a contiguous array of real COM VARIANTs for the DISPPARAMS struct used by IDispatch::Invoke
        /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221416(v=vs.85).aspx
        /// </summary>
        /// <param name="args">An array of arguments to wrap</param>
        /// <returns>DISPPARAMS structure ready to pass to IDispatch::Invoke</returns>
        private static ComTypes.DISPPARAMS PrepareDispatchArgs(object[] args)
        {
            var pDispParams = new ComTypes.DISPPARAMS();

            if ((args != null) && (args.Length != 0))
            {
                var variantStructSize = Marshal.SizeOf(typeof(VARIANT));
                pDispParams.cArgs = args.Length;

                var argsVariantLength = variantStructSize * pDispParams.cArgs;
                var variantArgsArray = Marshal.AllocHGlobal(argsVariantLength);

                // In IDispatch::Invoke, arguments are passed in reverse order
                IntPtr variantArgsArrayOffset = variantArgsArray + argsVariantLength;
                foreach (var arg in args)
                {
                    variantArgsArrayOffset -= variantStructSize;
                    Marshal.GetNativeVariantForObject(arg, variantArgsArrayOffset);
                }
                pDispParams.rgvarg = variantArgsArray;
            }
            return pDispParams;
        }

        [DllImport("oleaut32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        static extern Int32 VariantClear(IntPtr pvarg);

        /// <summary>
        /// frees all unmanaged memory assoicated with a DISPPARAMS structure
        /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221416(v=vs.85).aspx
        /// </summary>
        /// <param name="pDispParams"></param>
        private static void UnprepareDispatchArgs(ComTypes.DISPPARAMS pDispParams)
        {
            if (pDispParams.rgvarg != IntPtr.Zero)
            {
                // free the array of COM VARIANTs
                var variantStructSize = Marshal.SizeOf(typeof(VARIANT));
                var variantArgsArrayOffset = pDispParams.rgvarg;
                int argIndex = 0;
                while (argIndex < pDispParams.cArgs)
                {
                    VariantClear(variantArgsArrayOffset);
                    variantArgsArrayOffset += variantStructSize;
                    argIndex++;
                }
                Marshal.FreeHGlobal(pDispParams.rgvarg);
            }
        }

        /// <summary>
        /// A basic helper for IDispatch::Invoke
        /// </summary>
        /// <param name="obj">The IDispatch object of which you want to invoke a member on</param>
        /// <param name="memberId">The dispatch ID of the member to invoke</param>
        /// <param name="invokeKind">See InvokeKind enumeration</param>
        /// <param name="args">Array of arguments to pass to the call, or null for no args</param>
        /// <remarks>TODO support DISPATCH_PROPERTYPUTREF (property-set) which requires special handling</remarks>
        /// <returns>An object representing the return value from the called routine</returns>
        public static object Invoke(IDispatch obj, int memberId, InvokeKind invokeKind, object[] args = null)
        {
            var pDispParams = PrepareDispatchArgs(args);
            var pExcepInfo = new ComTypes.EXCEPINFO();

            int hr = obj.Invoke(memberId, ref _guid_null, 0, (uint)invokeKind,
                                    ref pDispParams, out object pVarResult, ref pExcepInfo, out uint pErrArg);

            UnprepareDispatchArgs(pDispParams);

            if (ComHelper.HRESULT_FAILED(hr))
            {
                if ((hr == (int)KnownComHResults.DISP_E_EXCEPTION) && (ComHelper.HRESULT_FAILED(pExcepInfo.scode)))
                {
                    throw Marshal.GetExceptionForHR(pExcepInfo.scode);
                }
                throw Marshal.GetExceptionForHR(hr);
            }

            return pVarResult;
        }
    }
}
