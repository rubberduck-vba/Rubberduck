using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.Parsing.ComReflection
{
    //See https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/
    public class ComVariant
    {
        internal static readonly IDictionary<VarEnum, string> TypeNames = new Dictionary<VarEnum, string>
        {
            {VarEnum.VT_DISPATCH, "Object"},
            {VarEnum.VT_VOID, string.Empty},
            {VarEnum.VT_VARIANT, "Variant"},
            {VarEnum.VT_BLOB_OBJECT, "Object"},
            {VarEnum.VT_STORED_OBJECT, "Object"},
            {VarEnum.VT_STREAMED_OBJECT, "Object"},
            {VarEnum.VT_BOOL, "Boolean"},
            {VarEnum.VT_BSTR, "String"},
            {VarEnum.VT_LPSTR, "String"},
            {VarEnum.VT_LPWSTR, "String"},
            {VarEnum.VT_I1, "Variant"}, // no signed byte type in VBA
            {VarEnum.VT_UI1, "Byte"},
            {VarEnum.VT_I2, "Integer"},
            {VarEnum.VT_UI2, "Variant"}, // no unsigned integer type in VBA
            {VarEnum.VT_I4, "Long"},
            {VarEnum.VT_UI4, "Variant"}, // no unsigned long integer type in VBA
            {VarEnum.VT_I8, "Variant"}, // LongLong on 64-bit VBA
            {VarEnum.VT_UI8, "Variant"}, // no unsigned LongLong integer type in VBA
            {VarEnum.VT_INT, "Long"}, // same as I4
            {VarEnum.VT_UINT, "Variant"}, // same as UI4
            {VarEnum.VT_DATE, "Date"},
            {VarEnum.VT_CY, "Currency"},
            {VarEnum.VT_DECIMAL, "Currency"}, // best match?
            {VarEnum.VT_EMPTY, "Empty"},
            {VarEnum.VT_R4, "Single"},
            {VarEnum.VT_R8, "Double"},
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

        public VarEnum VariantType { get; private set; }
        public object Value { get; private set; }

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
    }
}
