using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.ComReflection
{
    //See https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/
    public class ComVariant : IEquatable<ComVariant>
    {
        private static readonly bool ProcessIs32Bit = IntPtr.Size == 4;

        internal static readonly IDictionary<VarEnum, string> TypeNames = new Dictionary<VarEnum, string>
        {
            {VarEnum.VT_DISPATCH, Tokens.Object},
            {VarEnum.VT_VOID, string.Empty},
            {VarEnum.VT_VARIANT, Tokens.Variant},
            {VarEnum.VT_UNKNOWN, Tokens.Object},
            {VarEnum.VT_BLOB_OBJECT, Tokens.Object},
            {VarEnum.VT_STORED_OBJECT, Tokens.Object},
            {VarEnum.VT_STREAMED_OBJECT, Tokens.Object},
            {VarEnum.VT_BOOL, Tokens.Boolean},
            {VarEnum.VT_BSTR, Tokens.String},
            {VarEnum.VT_LPSTR, Tokens.LongPtr},
            {VarEnum.VT_PTR, Tokens.LongPtr},
            {VarEnum.VT_LPWSTR, Tokens.LongPtr},
            {VarEnum.VT_I1, Tokens.Variant}, // no signed byte type in VBA
            {VarEnum.VT_UI1, Tokens.Byte},
            {VarEnum.VT_I2, Tokens.Integer},
            {VarEnum.VT_UI2, Tokens.Variant}, // no unsigned integer type in VBA
            {VarEnum.VT_I4, Tokens.Long},
            {VarEnum.VT_UI4, Tokens.Variant}, // no unsigned long integer type in VBA
            {VarEnum.VT_I8, ProcessIs32Bit ? Tokens.Variant : Tokens.LongLong}, // LongLong on 64-bit VBA
            {VarEnum.VT_UI8, Tokens.Variant}, // no unsigned LongLong integer type in VBA
            {VarEnum.VT_INT, Tokens.Long}, // same as I4
            {VarEnum.VT_UINT, Tokens.Variant}, // same as UI4
            {VarEnum.VT_DATE, Tokens.Date},
            {VarEnum.VT_CY, Tokens.Currency},
            {VarEnum.VT_DECIMAL, Tokens.Decimal},
            {VarEnum.VT_EMPTY, Tokens.Empty},
            {VarEnum.VT_R4, Tokens.Single},
            {VarEnum.VT_R8, Tokens.Double}
        };

        [StructLayout(LayoutKind.Explicit)]
        private struct Variant
        {
            [FieldOffset(0)]
            public readonly ushort vt;
            [FieldOffset(2)]
            private readonly ushort wReserved1;
            [FieldOffset(4)]
            private readonly ushort wReserved2;
            [FieldOffset(6)]
            private readonly ushort wReserved3;

            //union
            [FieldOffset(8)]
            public readonly sbyte sbVal;
            [FieldOffset(8)]
            public readonly byte bVal;
            [FieldOffset(8)]
            public readonly short sVal;
            [FieldOffset(8)]
            public readonly ushort usVal;
            [FieldOffset(8)]
            public readonly int iVal;
            [FieldOffset(8)]
            public readonly uint uiVal;
            [FieldOffset(8)]
            public readonly long lVal;
            [FieldOffset(8)]
            public readonly ulong ulVal;
            [FieldOffset(8)]
            public readonly float fltVal;
            [FieldOffset(8)]
            public readonly double dblVal;
            [FieldOffset(8)]
            public readonly bool boolVal;
            //end union

            public IntPtr ToPointer() => IntPtr.Size == 4 ? new IntPtr(iVal) : new IntPtr(lVal);
        }

        public VarEnum VariantType { get; }
        public object Value { get; }
        public string TypeName => TypeNames.TryGetValue(VariantType, out var typeName) ? typeName : "Object";

        public ComVariant(IntPtr variant)
        {
            var members = Marshal.PtrToStructure<Variant>(variant);
            VariantType = (VarEnum)members.vt;

            // Note that these are technically flags, but it they are VT_BYREF or VT_ARRAY, the data area will always be a pointer.
            // ReSharper disable once SwitchStatementMissingSomeCases
            switch ((VarEnum)members.vt)
            {
                case VarEnum.VT_VOID:
                case VarEnum.VT_EMPTY:
                case VarEnum.VT_NULL:
                case VarEnum.VT_ERROR:
                    Value = null;
                    break;
                case VarEnum.VT_I2:
                    Value = members.sVal;
                    break;
                case VarEnum.VT_HRESULT:
                case VarEnum.VT_I4:
                    Value = members.iVal;
                    break;
                case VarEnum.VT_R4:
                    Value = members.fltVal;
                    break;
                case VarEnum.VT_R8:
                    Value = members.dblVal;
                    break;
                case VarEnum.VT_CY:
                    Value = members.dblVal * 10000;
                    break;
                case VarEnum.VT_DATE:
                    Value = DateTime.FromOADate(members.dblVal);
                    break;
                case VarEnum.VT_BSTR:
                    Value = Marshal.PtrToStringBSTR(members.ToPointer());
                    break;
                case VarEnum.VT_DISPATCH:
                    break;
                case VarEnum.VT_BOOL:
                    Value = members.boolVal;
                    break;
                case VarEnum.VT_I1:
                    Value = members.sbVal;
                    break;
                case VarEnum.VT_UI1:
                    Value = members.bVal;
                    break;
                case VarEnum.VT_UI2:
                    Value = members.usVal;
                    break;
                case VarEnum.VT_UI4:
                    Value = members.uiVal;
                    break;
                case VarEnum.VT_I8:
                    Value = members.lVal;
                    break;
                case VarEnum.VT_UI8:
                    Value = members.ulVal;
                    break;
                case VarEnum.VT_INT:
                    Value = IntPtr.Size == 4 ? members.iVal : members.lVal;
                    break;
                case VarEnum.VT_UINT:
                    Value = IntPtr.Size == 4 ? members.uiVal : members.ulVal;
                    break;
                case VarEnum.VT_LPSTR:
                    Value = Marshal.PtrToStringAnsi(members.ToPointer());
                    break;
                case VarEnum.VT_LPWSTR:
                    Value = Marshal.PtrToStringUni(members.ToPointer());
                    break;
                default:
                    Value = members.ToPointer();
                    break;
            }
        }

        public override bool Equals(object obj)
        {
            var other = obj as ComVariant;
            return other != null ? Equals(other) : Value.Equals(obj);
        }

        public bool Equals(ComVariant other)
        {
            if (other is null)
            {
                return false;
            }
            if (ReferenceEquals(this, other))
            {
                return true;
            }
            return VariantType == other.VariantType && Equals(Value, other.Value);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((int)VariantType * 397) ^ (Value?.GetHashCode() ?? 0);
            }
        }
    }
}
