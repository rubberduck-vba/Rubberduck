using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComTypedElement
    {
        string Name { get; }
        int Index { get; }
        DeclarationType Type { get; }
        object Value { get; }
        bool IsReferenceType { get; }
        string TypeName { get; }
        bool IsEnumMember { get; }
        bool IsAlias { get; }
        bool IsArray { get; }
    }

    public abstract class ComTypedElement : IComTypedElement
    {
        private static readonly HashSet<TYPEKIND> ReferenceTypeKinds = new HashSet<TYPEKIND>
        {
            TYPEKIND.TKIND_DISPATCH,
            TYPEKIND.TKIND_COCLASS,
            TYPEKIND.TKIND_INTERFACE
        };

        public virtual string Name { get; }
        public int Index { get; }
        public abstract DeclarationType Type { get; }
        public virtual object Value { get; private set; }
        public virtual bool IsReferenceType { get; private set; }
        public virtual bool IsArray { get; private set; }

        private string _type = string.Empty;

        public virtual string TypeName
        {
            get
            {
                var typeName = _type;
                if (IsEnumMember && ComProject.KnownEnumerations.TryGetValue(EnumGuid, out ComEnumeration enumType))
                {
                    var member = enumType.Members.FirstOrDefault(m => m.Value == Value);
                    if (member != null)
                    {
                        typeName = member.TypeName;
                    }
                }
                else if (IsAlias && ComProject.KnownAliases.TryGetValue(AliasGuid, out ComAlias alias))
                {
                    typeName = alias.Name;
                }
                return IsArray ? $"{typeName}()" : typeName;
            }
        }

        protected Guid EnumGuid { get; private set; } = Guid.Empty;
        public virtual bool IsEnumMember => !EnumGuid.Equals(Guid.Empty);
        protected Guid AliasGuid { get; private set; } = Guid.Empty;
        public virtual bool IsAlias => !AliasGuid.Equals(Guid.Empty);

        protected ComTypedElement(ITypeInfo info, VARDESC varDesc, string name, int index, bool constant = false)
        {
            Name = name;
            Index = index;

            if (constant)
            {
                GetValue(varDesc.desc.lpvarValue);
            }
            else
            {
                GetFieldType(varDesc.elemdescVar.tdesc, info);
            }
        }

        protected ComTypedElement(ITypeInfo info, VARDESC varDesc)
        {
            var names = new string[1];
            info.GetNames(varDesc.memid, names, names.Length, out int count);
            Debug.Assert(count == 1);
            Name = names[0];
            GetValue(varDesc.desc.lpvarValue);
        }

        protected ComTypedElement(ITypeInfo info, ELEMDESC elemDesc, string name, int index) : this(info, elemDesc.tdesc)
        {
            Name = name;
            Index = index;
        }

        protected ComTypedElement(ITypeInfo info, TYPEDESC desc)
        {
            GetFieldType(desc, info);
        }

        private void GetValue(IntPtr lpvarValue)
        {
            var value = new ComVariant(lpvarValue);
            Value = value.Value;
            _type = value.TypeName;
            IsArray = value.VariantType.HasFlag(VarEnum.VT_ARRAY);
        }

        private void GetFieldType(TYPEDESC desc, ITypeInfo info)
        {
            if (Convert.ToBoolean(desc.vt & (int)VarEnum.VT_PTR))
            {
                // ReSharper disable once TailRecursiveCall  <-- A stack allocation is *not* the heavy part of this call, the marshal is. Micro optimization at best.
                GetFieldType(Marshal.PtrToStructure<TYPEDESC>(desc.lpValue), info);
            }
            else if (Convert.ToBoolean(desc.vt & (int)VarEnum.VT_USERDEFINED))
            {
                int href;
                unchecked
                {
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
                            EnumGuid = attribs.guid;
                        }
                        else if (attribs.typekind == TYPEKIND.TKIND_ALIAS)
                        {
                            AliasGuid = attribs.guid;
                        }
                        IsReferenceType = ReferenceTypeKinds.Contains(attribs.typekind);
                        _type = new ComDocumentation(refTypeInfo, -1).Name;
                    }
                }
                catch (COMException) { }
            }
            else if (Convert.ToBoolean(desc.vt & (int)VarEnum.VT_SAFEARRAY) ||
                     Convert.ToBoolean(desc.vt & (int)VarEnum.VT_CARRAY) ||
                     Convert.ToBoolean(desc.vt & (int)VarEnum.VT_ARRAY))
            {
                GetFieldType(Marshal.PtrToStructure<TYPEDESC>(desc.lpValue), info);
                IsArray = true;
            }
            else
            {
                _type = ComVariant.TypeNames.TryGetValue((VarEnum)desc.vt, out var typeName) ? typeName : "Object";
            }
        }
    }
}
