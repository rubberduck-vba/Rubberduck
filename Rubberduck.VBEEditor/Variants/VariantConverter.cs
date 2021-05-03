using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.Variants
{
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
        public static object ChangeType(object value, VARENUM vt)
        {
            return ChangeType(value, vt, null);
        }

        public static object ChangeType(object value, VARENUM vt, CultureInfo cultureInfo)
        {
            object result = null;
            var hr = cultureInfo == null
                ? VariantNativeMethods.VariantChangeType(ref result, ref value, VariantConversionFlags.NO_FLAGS, vt)
                : VariantNativeMethods.VariantChangeTypeEx(ref result, ref value, cultureInfo.LCID, VariantConversionFlags.NO_FLAGS, vt);
            if (HResult.Failed(hr))
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