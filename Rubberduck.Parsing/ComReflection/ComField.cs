using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComField
    {
        public string Name { get; }
        public int Index { get; }
        public DeclarationType Type { get; }
        public object DefaultValue { get; set; }
        public string ValueType { get; set; }
        public VARFLAGS Flags { get; set; }

        public ComField(string name, VARDESC varDesc, int index, DeclarationType type)
        {
            Name = name;
            Index = index;
            Type = type;

            Flags = (VARFLAGS)varDesc.wVarFlags;

            if (Type == DeclarationType.Constant)
            {
                var value = new ComVariant(varDesc.desc.lpvarValue);
                DefaultValue = value.Value;
                ValueType = ComVariant.TypeNames.TryGetValue(value.VariantType, out var typeName) ? typeName : "Object";
            }
            else
            {
                Debug.Assert(varDesc.varkind == VARKIND.VAR_PERINSTANCE);
                ValueType = ComVariant.TypeNames.TryGetValue((VarEnum)varDesc.elemdescVar.tdesc.vt, out var typeName) ? typeName : "Object";                
            }
        }
    }
}
