using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComField : ComTypedElement
    {
        public override DeclarationType Type { get; }
        public VARFLAGS Flags { get; }

        public ComField(ITypeInfo info, VARDESC varDesc, string name, int index, DeclarationType type) : base(info, varDesc, name, index, type == DeclarationType.Constant)
        {
            Type = type;
            Flags = (VARFLAGS)varDesc.wVarFlags;
        }

        public ComField(ITypeInfo info, TYPEDESC desc) : base(info, desc) { }
    }
}
