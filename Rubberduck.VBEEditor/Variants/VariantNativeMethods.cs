using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.Variants
{
    public enum VARENUM
    {
        VT_EMPTY = 0x0000,
        VT_NULL = 0x0001,
        VT_I2 = 0x0002,
        VT_I4 = 0x0003,
        VT_R4 = 0x0004,
        VT_R8 = 0x0005,
        VT_CY = 0x0006,
        VT_DATE = 0x0007,
        VT_BSTR = 0x0008,
        VT_DISPATCH = 0x0009,
        VT_ERROR = 0x000A,
        VT_BOOL = 0x000B,
        VT_VARIANT = 0x000C,
        VT_UNKNOWN = 0x000D,
        VT_DECIMAL = 0x000E,
        VT_I1 = 0x0010,
        VT_UI1 = 0x0011,
        VT_UI2 = 0x0012,
        VT_UI4 = 0x0013,
        VT_I8 = 0x0014,
        VT_UI8 = 0x0015,
        VT_INT = 0x0016,
        VT_UINT = 0x0017,
        VT_VOID = 0x0018,
        VT_HRESULT = 0x0019,
        VT_PTR = 0x001A,
        VT_SAFEARRAY = 0x001B,
        VT_CARRAY = 0x001C,
        VT_USERDEFINED = 0x001D,
        VT_LPSTR = 0x001E,
        VT_LPWSTR = 0x001F,
        VT_RECORD = 0x0024,
        VT_INT_PTR = 0x0025,
        VT_UINT_PTR = 0x0026,
        VT_ARRAY = 0x2000,
        VT_BYREF = 0x4000
    }

    [Flags]
    public enum VariantConversionFlags : ushort
    {
        NO_FLAGS = 0x00,
        VARIANT_NOVALUEPROP = 0x01,     //Prevents the function from attempting to coerce an object to a fundamental type by getting the Value property. Applications should set this flag only if necessary, because it makes their behavior inconsistent with other applications.
        VARIANT_ALPHABOOL = 0x02,       //Converts a VT_BOOL value to a string containing either "True" or "False".
        VARIANT_NOUSEROVERRIDE = 0x04,  //For conversions to or from VT_BSTR, passes LOCALE_NOUSEROVERRIDE to the core coercion routines.
        VARIANT_LOCALBOOL = 0x08        //For conversions from VT_BOOL to VT_BSTR and back, uses the language specified by the locale in use on the local computer. 
    }

    [Flags]
    public enum VariantComparisonFlags : ulong
    {
        NO_FLAGS = 0x00,
        NORM_IGNORECASE = 0x00000001,       //Ignore case.
        NORM_IGNORENONSPACE = 0x00000002,   //Ignore nonspace characters.
        NORM_IGNORESYMBOLS = 0x00000004,    //Ignore symbols.
        NORM_IGNOREWIDTH = 0x00000008,      //Ignore string width.
        NORM_IGNOREKANATYPE = 0x00000040,   //Ignore Kana type.
        NORM_IGNOREKASHIDA = 0x00040000     //Ignore Arabic kashida characters. 
    }

    public enum VariantComparisonResults : int
    {
        VARCMP_LT = 0,      //pvarLeft is less than pvarRight.
        VARCMP_EQ = 1,      //The parameters are equal.
        VARCMP_GT = 2,      //pvarLeft is greater than pvarRight.
        VARCMP_NULL = 3     //Either expression is NULL.
    }

    public static class HResult
    {
        internal static bool Succeeded(int hr) => hr >= 0;
        internal static bool Failed(int hr) => hr < 0;
    }

    internal static class VariantNativeMethods
    {
        private const string dllName = "oleaut32.dll";

        // HRESULT VariantChangeType(
        //   VARIANTARG       *pvargDest,
        //   const VARIANTARG *pvarSrc,
        //   USHORT           wFlags,
        //   VARTYPE          vt
        // );
        [DllImport(dllName, EntryPoint = "VariantChangeType", CharSet = CharSet.Auto, SetLastError = true, PreserveSig = true)]
        internal static extern int VariantChangeType(ref object pvargDest, ref object pvarSrc, VariantConversionFlags wFlags, VARENUM vt);

        // HRESULT VariantChangeTypeEx(
        //   VARIANTARG        *pvargDest,
        //   const VARIANTARG  *pvarSrc,
        //   LCID              lcid,
        //   USHORT            wFlags,
        //   VARTYPE           vt
        // );
        [DllImport(dllName, EntryPoint = "VariantChangeTypeEx", CharSet = CharSet.Auto, SetLastError = true, PreserveSig = true)]
        internal static extern int VariantChangeTypeEx(ref object pvargDest, ref object pvarSrc, int lcid, VariantConversionFlags wFlags, VARENUM vt);

        // HRESULT VarCmp(
        //   LPVARIANT pvarLeft,
        //   LPVARIANT pvarRight,
        //   LCID lcid,
        //   ULONG dwFlags
        // );
        [DllImport(dllName, EntryPoint = "VarCmp", CharSet = CharSet.Auto, SetLastError = true, PreserveSig = true)]
        internal static extern int VarCmp(ref object pvarLeft, ref object pvarRight, int lcid, ulong dwFlags);
    }
}
