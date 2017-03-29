using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.ComReflection
{
    //See https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/
    public class ComVariant : IEquatable<ComVariant>
    {
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
            {VarEnum.VT_LPWSTR, Tokens.LongPtr},
            {VarEnum.VT_I1, Tokens.Variant}, // no signed byte type in VBA
            {VarEnum.VT_UI1, Tokens.Byte},
            {VarEnum.VT_I2, Tokens.Integer},
            {VarEnum.VT_UI2, Tokens.Variant}, // no unsigned integer type in VBA
            {VarEnum.VT_I4, Tokens.Long},
            {VarEnum.VT_UI4, Tokens.Variant}, // no unsigned long integer type in VBA
            {VarEnum.VT_I8, Tokens.Variant}, // LongLong on 64-bit VBA
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


        [StructLayout(LayoutKind.Sequential)]
        private struct Variant
        {
            public readonly ushort vt;
            private readonly ushort wReserved1;
            private readonly ushort wReserved2;
            private readonly ushort wReserved3;
            private readonly int data01;
            private readonly int data02;
        }

        public VarEnum VariantType { get; }
        public object Value { get; }

        public ComVariant(IntPtr variant)
        {
            Value = Marshal.GetObjectForNativeVariant(variant);
            var members = (Variant)Marshal.PtrToStructure(variant, typeof(Variant));
            VariantType = (VarEnum)members.vt;
            if (Value == null && VariantType == VarEnum.VT_BSTR)
            {
                Value = string.Empty;
            }
        }

        public override bool Equals(object obj)
        {
            var other = obj as ComVariant;
            return other != null ? Equals(other) : Value.Equals(obj);
        }

        public bool Equals(ComVariant other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
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
