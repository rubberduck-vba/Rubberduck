using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name}")]
    public class ComField
    {
        public string Name { get; set; }
        public int Index { get; set; }
        public bool IsConstant { get; set; }
        public object DefaultValue { get; set; }
        public VarEnum DefaultValueType { get; set; }
        public VARFLAGS Flags { get; set; }

        public ComField(ITypeInfo info, VARDESC varDesc, int index)
        {
            Index = index;

            var names = new string[255];
            int length;
            info.GetNames(varDesc.memid, names, 255, out length);
            Debug.Assert(length >= 1);
            Name = names[0];

            IsConstant = varDesc.varkind.HasFlag(VARKIND.VAR_CONST);
            Flags = (VARFLAGS)varDesc.wVarFlags;

            if (IsConstant)
            {
                var value = new ComVariant(varDesc.desc.lpvarValue);
                DefaultValue = value.Value;
                DefaultValueType = value.VariantType;
            }
        }
    }
}
