using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name} = {Value} ({ValueType})")]
    public class ComEnumerationMember
    {
        public string Name { get; private set; }
        public int Value { get; private set; }
        public VarEnum ValueType { get; private set; }

        public ComEnumerationMember(ITypeInfo info, VARDESC varDesc)
        {
            var value = new ComVariant(varDesc.desc.lpvarValue);
            Value = (int)value.Value;
            ValueType = value.VariantType;

            var names = new string[255];
            int count;
            info.GetNames(varDesc.memid, names, 1, out count);
            Debug.Assert(count == 1);
            Name = names[0];
        }
    }
}
