using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name}")]
    public class ComField
    {
        public string Name { get; }
        public int Index { get; }
        public DeclarationType Type { get; }
        public object DefaultValue { get; }

        private string _valueType = "Object";
        public string ValueType => IsArray ? $"{_valueType}()" : _valueType;

        private Guid _enumGuid = Guid.Empty;
        public bool IsEnumMember => !_enumGuid.Equals(Guid.Empty);

        public bool IsArray { get; private set; }
        public VARFLAGS Flags { get; }

        public ComField(ITypeInfo info, string name, VARDESC varDesc, int index, DeclarationType type)
        {
            Name = name;
            Index = index;
            Type = type;

            Flags = (VARFLAGS)varDesc.wVarFlags;

            if (Type == DeclarationType.Constant)
            {
                var value = new ComVariant(varDesc.desc.lpvarValue);
                DefaultValue = value.Value;

                if (ComVariant.TypeNames.TryGetValue(value.VariantType, out string typeName))
                {
                    _valueType = typeName;
                }

                if (value.VariantType.HasFlag(VarEnum.VT_ARRAY))
                {
                    IsArray = true;
                }
            }
            else
            {
                GetFieldType(varDesc.elemdescVar.tdesc, info);
                if (!IsEnumMember || !ComProject.KnownEnumerations.TryGetValue(_enumGuid, out ComEnumeration enumType))
                {
                    return;
                }
                var member = enumType.Members.FirstOrDefault(m => m.Value == (int)DefaultValue);
                _valueType = member != null ? member.Name : _valueType;
            }
        }

        private void GetFieldType(TYPEDESC desc, ITypeInfo info)
        {
            var vt = (VarEnum)desc.vt;
            TYPEDESC tdesc;

            switch (vt)
            {
                case VarEnum.VT_PTR:
                    tdesc = (TYPEDESC)Marshal.PtrToStructure(desc.lpValue, typeof(TYPEDESC));
                    GetFieldType(tdesc, info);
                    break;
                case VarEnum.VT_USERDEFINED:
                    int href;
                    unchecked
                    {
                        href = (int)(desc.lpValue.ToInt64() & 0xFFFFFFFF);
                    }
                    try
                    {
                        info.GetRefTypeInfo(href, out ITypeInfo refTypeInfo);

                        refTypeInfo.GetTypeAttr(out IntPtr attribPtr);
                        var attribs = (TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(TYPEATTR));
                        if (attribs.typekind == TYPEKIND.TKIND_ENUM)
                        {
                            _enumGuid = attribs.guid;
                        }
                        _valueType = new ComDocumentation(refTypeInfo, -1).Name;
                        refTypeInfo.ReleaseTypeAttr(attribPtr);
                    }
                    catch (COMException) { }
                    break;
                case VarEnum.VT_SAFEARRAY:
                case VarEnum.VT_CARRAY:
                case VarEnum.VT_ARRAY:
                    tdesc = (TYPEDESC)Marshal.PtrToStructure(desc.lpValue, typeof(TYPEDESC));
                    GetFieldType(tdesc, info);
                    IsArray = true;
                    break;
                default:
                    if (ComVariant.TypeNames.TryGetValue(vt, out string result))
                    {
                        _valueType = result;
                    }
                    break;
            }
        }
    }
}
