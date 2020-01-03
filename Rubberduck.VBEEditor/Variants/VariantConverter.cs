using System;
using System.Globalization;
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

    /// <summary>
    /// Handles variant conversions, enabling us to have same implicit conversion behaviors within
    /// .NET as we can observe it from VBA/VB6.
    /// </summary>
    /// <remarks>
    /// The <see cref="VariantChangeType"/> function is the same one used internally by VBA/VB6.
    /// However, we have to wrap the metadata, which the class helps with.
    /// 
    /// See the link for details on how marshaling are handled with <see cref="object"/>
    /// https://docs.microsoft.com/en-us/dotnet/framework/interop/default-marshaling-for-objects
    /// </remarks>
    public static class VariantConverter
    {
        private const string dllName = "oleaut32.dll";

        // HRESULT VariantChangeType(
        //   VARIANTARG       *pvargDest,
        //   const VARIANTARG *pvarSrc,
        //   USHORT           wFlags,
        //   VARTYPE          vt
        // );
        [DllImport(dllName, EntryPoint = "VariantChangeType", CharSet = CharSet.Auto, SetLastError = true, PreserveSig = true)]
        private static extern int VariantChangeType(ref object pvargDest, ref object pvarSrc, VariantConversionFlags wFlags, VARENUM vt);

        // HRESULT VariantChangeTypeEx(
        //   VARIANTARG        *pvargDest,
        //   const VARIANTARG  *pvarSrc,
        //   LCID              lcid,
        //   USHORT            wFlags,
        //   VARTYPE           vt
        // );
        [DllImport(dllName, EntryPoint = "VariantChangeTypeEx", CharSet = CharSet.Auto, SetLastError = true, PreserveSig = true)]
        private static extern int VariantChangeTypeEx(ref object pvargDest, ref object pvarSrc, int lcid, VariantConversionFlags wFlags, VARENUM vt);

        public static object ChangeType(object value, VARENUM vt)
        {
            return ChangeType(value, vt, null);
        }

        private static bool HRESULT_FAILED(int hr) => hr < 0;
        public static object ChangeType(object value, VARENUM vt, CultureInfo cultureInfo)
        {
            object result = null;
            var hr = cultureInfo == null
                ? VariantChangeType(ref result, ref value, VariantConversionFlags.NO_FLAGS, vt)
                : VariantChangeTypeEx(ref result, ref value, cultureInfo.LCID, VariantConversionFlags.NO_FLAGS, vt);
            if (HRESULT_FAILED(hr))
            {
                throw Marshal.GetExceptionForHR(hr);
            }

            return result;
        }

        public static object ChangeType(object value, Type targetType)
        {
            return ChangeType(value, GetVarEnum(targetType));
        }

        public static object ChangeType(object value, Type targetType, CultureInfo culture)
        {
            return ChangeType(value, GetVarEnum(targetType), culture);
        }

        public static VARENUM GetVarEnum(Type target)
        {
            switch (target)
            {
                case null:
                    return VARENUM.VT_EMPTY;
                case Type dbNull when dbNull == typeof(DBNull):
                    return VARENUM.VT_NULL;
                case Type err when err == typeof(ErrorWrapper):
                    return VARENUM.VT_ERROR;
                case Type disp when disp == typeof(DispatchWrapper):
                    return VARENUM.VT_DISPATCH;
                case Type unk when unk == typeof(UnknownWrapper):
                    return VARENUM.VT_UNKNOWN;
                case Type cy when cy == typeof(CurrencyWrapper):
                    return VARENUM.VT_CY;
                case Type b when b == typeof(bool):
                    return VARENUM.VT_BOOL;
                case Type s when s == typeof(sbyte):
                    return VARENUM.VT_I1;
                case Type b when b == typeof(byte):
                    return VARENUM.VT_UI1;
                case Type i16 when i16 == typeof(short):
                    return VARENUM.VT_I2;
                case Type ui16 when ui16 == typeof(ushort):
                    return VARENUM.VT_UI2;
                case Type i32 when i32 == typeof(int):
                    return VARENUM.VT_I4;
                case Type ui32 when ui32 == typeof(uint):
                    return VARENUM.VT_UI4;
                case Type i64 when i64 == typeof(long):
                    return VARENUM.VT_I8;
                case Type ui64 when ui64 == typeof(ulong):
                    return VARENUM.VT_UI8;
                case Type sng when sng == typeof(float):
                    return VARENUM.VT_R4;
                case Type dbl when dbl == typeof(double):
                    return VARENUM.VT_R8;
                case Type dec when dec == typeof(decimal):
                    return VARENUM.VT_DECIMAL;
                case Type dt when dt == typeof(DateTime):
                    return VARENUM.VT_DATE;
                case Type s when s == typeof(string):
                    return VARENUM.VT_BSTR;
                //case Type a when a == typeof(Array):
                //    return VARENUM.VT_ARRAY;
                case Type obj when obj == typeof(object):
                case Type var when var == typeof(VariantWrapper):
                    return VARENUM.VT_VARIANT;
                default:
                    throw new NotSupportedException("Unrecognized system type that cannot be mapped to a VARENUM out of the box.");
            }
        }

        public static VARENUM GetVarEnum(TypeCode typeCode)
        {
            switch (typeCode)
            {
                case TypeCode.Empty:
                    return VARENUM.VT_EMPTY;
                case TypeCode.Object:
                    return VARENUM.VT_UNKNOWN;
                case TypeCode.DBNull:
                    return VARENUM.VT_NULL;
                case TypeCode.Boolean:
                    return VARENUM.VT_BOOL;
                case TypeCode.Char:
                    return VARENUM.VT_UI2;
                case TypeCode.SByte:
                    return VARENUM.VT_I1;
                case TypeCode.Byte:
                    return VARENUM.VT_UI1;
                case TypeCode.Int16:
                    return VARENUM.VT_I2;
                case TypeCode.UInt16:
                    return VARENUM.VT_UI2;
                case TypeCode.Int32:
                    return VARENUM.VT_I4;
                case TypeCode.UInt32:
                    return VARENUM.VT_UI4;
                case TypeCode.Int64:
                    return VARENUM.VT_I8;
                case TypeCode.UInt64:
                    return VARENUM.VT_UI8;
                case TypeCode.Single:
                    return VARENUM.VT_R4;
                case TypeCode.Double:
                    return VARENUM.VT_R8;
                case TypeCode.Decimal:
                    return VARENUM.VT_DECIMAL;
                case TypeCode.DateTime:
                    return VARENUM.VT_DATE;
                case TypeCode.String:
                    return VARENUM.VT_BSTR;
                default:
                    throw new ArgumentOutOfRangeException(nameof(typeCode), typeCode, null);
            }
        }
    }
}