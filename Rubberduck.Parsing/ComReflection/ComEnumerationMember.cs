using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "} = {" + nameof(Value) + "} ({" + nameof(ValueType) + "})")]
    public class ComEnumerationMember
    {
        public string Name { get; }
        public int Value { get; }
        public VarEnum ValueType { get; }

        public ComEnumerationMember(ITypeInfo info, VARDESC varDesc)
        {
            var value = new ComVariant(varDesc.desc.lpvarValue);
            Value = (int)value.Value;
            ValueType = value.VariantType;

            var names = new string[255];
            info.GetNames(varDesc.memid, names, 1, out var count);
            Debug.Assert(count == 1);
            Name = names[0];
        }
    }
}
