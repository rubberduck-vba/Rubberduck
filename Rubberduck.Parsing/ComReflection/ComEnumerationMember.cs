using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name} = {Value} ({ValueType})")]
    public class ComEnumerationMember
    {
        public string Name { get; }
        public int Value { get; }
        public VarEnum ValueType { get; }
        ComEnumeration Parent { get; }
        public ComProject Project => Parent?.Project;

        public ComEnumerationMember(ComEnumeration parent, ITypeInfo info, VARDESC varDesc)
        {
            Parent = parent;

            var value = new ComVariant(varDesc.desc.lpvarValue);
            Value = (int)value.Value;
            ValueType = value.VariantType;

            var names = new string[1];
            info.GetNames(varDesc.memid, names, names.Length, out int count);
            Debug.Assert(count == 1);
            Name = names[0];
        }
    }
}
