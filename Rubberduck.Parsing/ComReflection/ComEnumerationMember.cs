using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComEnumerationMember : ComTypedElement
    {
        public override DeclarationType Type => DeclarationType.EnumerationMember;
        public override string TypeName => "Integer";
        public override bool IsEnumMember => false;
        public override bool IsAlias => false;
        public override bool IsReferenceType => false;
        public override bool IsArray => false;

        public ComEnumerationMember(ITypeInfo info, VARDESC varDesc) : base(info, varDesc) { }
    }
}