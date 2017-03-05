using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{DeclarationName}")]
    public class ComParameter
    {
        public string Name { get; private set; }

        public string DeclarationName
        {
            get
            {
                return string.Format("{0}{1} {2} As {3}{4}{5}",
                    IsOptional ? "Optional " : string.Empty,
                    IsByRef ? "ByRef" : "ByVal",
                    Name,
                    TypeName,
                    IsOptional && DefaultValue != null ? " = " : string.Empty,
                    IsOptional && DefaultValue != null ? 
                        IsEnumMember ? DefaultAsEnum : DefaultValue 
                        : string.Empty);
            }
        }

        public bool IsArray { get; private set; }
        public bool IsByRef { get; private set; }
        public bool IsOptional { get; private set; }
        public bool IsParamArray { get; set; }

        private Guid _enumGuid = Guid.Empty;
        public bool IsEnumMember
        {
            get { return !_enumGuid.Equals(Guid.Empty); }
        }
        public object DefaultValue { get; private set; }
        public string DefaultAsEnum { get; private set; }

        private string _type = "Object";
        public string TypeName
        {
            get
            {
                return IsArray ? _type + "()" : _type;
            }
        }

        public ComParameter(ELEMDESC elemDesc, ITypeInfo info, string name)
        {
            Name = name;
            var paramDesc = elemDesc.desc.paramdesc;
            GetParameterType(elemDesc.tdesc, info);
            IsOptional = paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT);
            if (!paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FHASDEFAULT) || string.IsNullOrEmpty(name))
            {
                DefaultAsEnum = string.Empty;
                return;
            }

            //lpVarValue points to a PARAMDESCEX structure, but we don't care about the cBytes here at all. 
            //Offset and dereference the VARIANTARG directly.
            var defValue = new ComVariant(paramDesc.lpVarValue + Marshal.SizeOf(typeof(ulong)));
            DefaultValue = defValue.Value;

            ComEnumeration enumType;
            if (!IsEnumMember || !ComProject.KnownEnumerations.TryGetValue(_enumGuid, out enumType))
            {
                return;
            }
            var member = enumType.Members.FirstOrDefault(m => m.Value == (int)DefaultValue);
            DefaultAsEnum = member != null ? member.Name : string.Empty;
        }

        //This overload should only be used for retrieving the TypeName from a random TYPEATTR. TODO: Should be a base class of ComParameter instead.
        public ComParameter(TYPEATTR attributes, ITypeInfo info)
        {
            GetParameterType(attributes.tdescAlias, info);
        }

        private void GetParameterType(TYPEDESC desc, ITypeInfo info)
        {
            var vt = (VarEnum)desc.vt;
            TYPEDESC tdesc;

            switch (vt)
            {
                case VarEnum.VT_PTR:
                    tdesc = (TYPEDESC)Marshal.PtrToStructure(desc.lpValue, typeof(TYPEDESC));
                    GetParameterType(tdesc, info);
                    IsByRef = true;                  
                    break;
                case VarEnum.VT_USERDEFINED:
                    int href;
                    unchecked
                    {
                        href = (int)(desc.lpValue.ToInt64() & 0xFFFFFFFF);
                    }
                    try
                    {
                        ITypeInfo refTypeInfo;
                        info.GetRefTypeInfo(href, out refTypeInfo);

                        IntPtr attribPtr;
                        refTypeInfo.GetTypeAttr(out attribPtr);
                        var attribs = (TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(TYPEATTR));
                        if (attribs.typekind == TYPEKIND.TKIND_ENUM)
                        {
                            _enumGuid = attribs.guid;
                        }
                        _type = new ComDocumentation(refTypeInfo, -1).Name;
                        refTypeInfo.ReleaseTypeAttr(attribPtr);
                    }                    
                    catch (COMException) { }
                    break;
                case VarEnum.VT_SAFEARRAY:
                case VarEnum.VT_CARRAY:
                case VarEnum.VT_ARRAY:
                    tdesc = (TYPEDESC)Marshal.PtrToStructure(desc.lpValue, typeof(TYPEDESC));
                    GetParameterType(tdesc, info);
                    IsArray = true;
                    break;
                default:
                    string result;
                    if (ComVariant.TypeNames.TryGetValue(vt, out result))
                    {
                        _type = result;
                    }
                    break;
            }
        }
    }
}
