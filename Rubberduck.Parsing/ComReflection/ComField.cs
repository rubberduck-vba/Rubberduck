using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComField
    {
        [DataMember(IsRequired = true)]
        public string Name { get; private set; }

        [DataMember(IsRequired = true)]
        public int Index { get; private set; }

        [DataMember(IsRequired = true)]
        public DeclarationType Type { get; private set; }

        [DataMember(IsRequired = true)]
        public object DefaultValue { get; private set; }

        [DataMember(IsRequired = true)]
        public bool IsReferenceType { get; private set; }

        [DataMember(IsRequired = true)]
        private string _valueType = Tokens.Object;
        public string ValueType => _valueType;

        [DataMember(IsRequired = true)]
        private Guid _enumGuid = Guid.Empty;
        public bool IsEnumMember => !_enumGuid.Equals(Guid.Empty);

        [DataMember(IsRequired = true)]
        public bool IsArray { get; private set; }
        [DataMember(IsRequired = true)]
        public VARFLAGS Flags { get; private set; }

        [DataMember(IsRequired = true)]
        IComBase Parent { get; set; }
        public ComProject Project => Parent?.Project;

        public ComField(IComBase parent, ITypeInfo info, string name, VARDESC varDesc, int index, DeclarationType type)
        {
            Parent = parent;
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
            }
        }

        private static readonly HashSet<TYPEKIND> ReferenceTypeKinds = new HashSet<TYPEKIND>
        {
            TYPEKIND.TKIND_DISPATCH,
            TYPEKIND.TKIND_COCLASS,
            TYPEKIND.TKIND_INTERFACE
        };

        private void GetFieldType(TYPEDESC desc, ITypeInfo info)
        {
            var vt = (VarEnum)desc.vt;
            TYPEDESC tdesc;

            if (vt == VarEnum.VT_PTR)
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
                GetFieldType(tdesc, info);
            }
            else if (vt == VarEnum.VT_USERDEFINED)
            {
                int href;
                unchecked
                {
                    //The href is a long, but the size of lpValue depends on the platform, so truncate it after the lword.
                    href = (int)(desc.lpValue.ToInt64() & 0xFFFFFFFF);
                }
                try
                {
                    info.GetRefTypeInfo(href, out ITypeInfo refTypeInfo);
                    refTypeInfo.GetTypeAttr(out IntPtr attribPtr);
                    using (DisposalActionContainer.Create(attribPtr, refTypeInfo.ReleaseTypeAttr))
                    {
                        var attribs = Marshal.PtrToStructure<TYPEATTR>(attribPtr);
                        if (attribs.typekind == TYPEKIND.TKIND_ENUM)
                        {
                            _enumGuid = attribs.guid;
                        }
                        IsReferenceType = ReferenceTypeKinds.Contains(attribs.typekind);
                        _valueType = new ComDocumentation(refTypeInfo, ComDocumentation.LibraryIndex).Name;
                    }
                }
                catch (COMException) { }
            }
            else if (vt == VarEnum.VT_SAFEARRAY || vt == VarEnum.VT_CARRAY || vt.HasFlag(VarEnum.VT_ARRAY))
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
                GetFieldType(tdesc, info);
                IsArray = true;
            }
            else
            {
                if (ComVariant.TypeNames.TryGetValue(vt, out string result))
                {
                    _valueType = result;
                }
            }
        }
    }
}
