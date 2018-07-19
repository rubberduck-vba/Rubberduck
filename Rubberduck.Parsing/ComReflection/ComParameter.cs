using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComParameter : ComTypedElement
    {
        public override string Name => base.Name ?? $"{Index}unnamedParameter";
        public override DeclarationType Type => DeclarationType.Parameter;
        public bool IsByRef { get; }
        public bool IsOptional { get; }
        public bool IsReturnValue { get; }
        public bool IsParamArray { get; set; }
        public object DefaultValue { get; }

        public string DefaultAsEnumMember
        {
            get
            {
                if (!IsEnumMember || !ComProject.KnownEnumerations.TryGetValue(EnumGuid, out ComEnumeration enumType))
                {
                    return string.Empty;
                }
                var member = enumType.Members.FirstOrDefault(m => m.Value == DefaultValue);
                return member != null ? member.Name : string.Empty;
            }
        }

        public ComParameter(ITypeInfo info, ELEMDESC elemDesc, string name, int index) : base(info, elemDesc, name, index)
        {
            var paramDesc = elemDesc.desc.paramdesc;
            IsOptional = paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT);
            IsReturnValue = paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FRETVAL);
            IsByRef = (VarEnum) elemDesc.tdesc.vt == VarEnum.VT_PTR;

            if (!paramDesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FHASDEFAULT) || string.IsNullOrEmpty(name))
            {
                return;
            }

            //lpVarValue points to a PARAMDESCEX structure, but we don't care about the cBytes here at all. 
            //Offset and dereference the VARIANTARG directly.
            var defValue = new ComVariant(paramDesc.lpVarValue + Marshal.SizeOf(typeof(ulong)));
            DefaultValue = defValue.Value;
        }
    }
}
