using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.Utility;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;

namespace Rubberduck.Parsing.ComReflection
{
#if DEBUG
    [DebuggerDisplay("{" + nameof(DeclarationName) + "}")]
#endif
    public class ComParameter
    {
        public static ComParameter Void = new ComParameter { _typeName = new ComTypeName(null, string.Empty) };

        public string Name { get; }

#if DEBUG
        public string DeclarationName => string.Format("{0}{1} {2} As {3}{4}{5}",
            IsOptional ? "Optional " : string.Empty,
            IsByRef ? "ByRef" : "ByVal",
            Name,
            TypeName,
            IsOptional && DefaultValue != null ? " = " : string.Empty,
            IsOptional && DefaultValue != null ?
                _typeName.IsEnumMember ? DefaultAsEnum : DefaultValue
                : string.Empty);
#endif

        public bool IsArray { get; private set; }
        public bool IsByRef { get; private set; }
        public bool IsOptional { get; }
        public bool IsReturnValue { get; }
        public bool IsParamArray { get; set; }
        public object DefaultValue { get; }

        public string DefaultAsEnum
        {
            get
            {
                if (!_typeName.IsEnumMember || !ComProject.KnownEnumerations.TryGetValue(_typeName.EnumGuid, out ComEnumeration enumType))
                {
                    return string.Empty;
                }
                var member = enumType.Members.FirstOrDefault(m => m.Value == (int)DefaultValue);
                return member != null ? member.Name : string.Empty;
            }
        }

        private ComTypeName _typeName;
        public string TypeName => IsArray ? $"{_typeName.Name}()" : _typeName.Name;

        ComMember Parent { get; }
        public ComProject Project => Parent?.Project;

        private ComParameter() { }

        public ComParameter(ComMember parent, ELEMDESC elemDesc, ITypeInfo info, string name)
        {
            Debug.Assert(name != null, "Parameter name is null");

            Parent = parent;
            Name = name;
            var paramDesc = elemDesc.desc.paramdesc;
            GetParameterType(elemDesc.tdesc, info);
            IsOptional = paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT);
            IsReturnValue = paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FRETVAL);
            if (!paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FHASDEFAULT) || string.IsNullOrEmpty(name))
            {
                return;
            }

            //lpVarValue points to a PARAMDESCEX structure, but we don't care about the cBytes here at all. 
            //Offset and dereference the VARIANTARG directly.
            var defValue = new ComVariant(paramDesc.lpVarValue + Marshal.SizeOf(typeof(ulong)));
            DefaultValue = defValue.Value;
        }

        //This overload should only be used for retrieving the TypeName from a random TYPEATTR. TODO: This really belongs somewhere else.
        public ComParameter(TYPEATTR attributes, ITypeInfo info)
        {
            GetParameterType(attributes.tdescAlias, info);
        }

        private void GetParameterType(TYPEDESC desc, ITypeInfo info)
        {
            var vt = (VarEnum)desc.vt;
            TYPEDESC tdesc;

            if (vt == VarEnum.VT_PTR)
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
                GetParameterType(tdesc, info);
                IsByRef = true;
            }
            else if (vt == VarEnum.VT_USERDEFINED)
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
                        var type = new ComDocumentation(refTypeInfo, -1).Name;
                        if (attribs.typekind == TYPEKIND.TKIND_ENUM)
                        {
                            _typeName = new ComTypeName(Project, type, attribs.guid, Guid.Empty);
                        }
                        else if (attribs.typekind == TYPEKIND.TKIND_ALIAS)
                        {
                            _typeName = new ComTypeName(Project, type, Guid.Empty, attribs.guid);
                        }
                        else
                        {
                            _typeName = new ComTypeName(Project, type);
                        }
                    }
                }
                catch (COMException)
                {
                    _typeName = new ComTypeName(Project, "Object");
                }
            }
            else if (vt == VarEnum.VT_SAFEARRAY || vt == VarEnum.VT_CARRAY || vt.HasFlag(VarEnum.VT_ARRAY))
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
                GetParameterType(tdesc, info);
                IsArray = true;
            }
            else if (vt == VarEnum.VT_HRESULT)
            {
                _typeName = new ComTypeName(Project, "Long");
            }
            else
            {
                _typeName = new ComTypeName(Project, (ComVariant.TypeNames.TryGetValue(vt, out string result)) ? result : "Object");
            }
        }
    }
}
