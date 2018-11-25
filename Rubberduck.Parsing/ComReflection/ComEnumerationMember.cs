using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [KnownType(typeof(ComEnumeration))]
    [DebuggerDisplay("{Name} = {Value} ({ValueType})")]
    public class ComEnumerationMember
    {
        [DataMember(IsRequired = true)]
        public string Name { get; private set; }

        [DataMember(IsRequired = true)]
        public int Value { get; private set; }

        [DataMember(IsRequired = true)]
        public VarEnum ValueType { get; private set; }

        [DataMember(IsRequired = true)]
        ComEnumeration Parent { get; set; }

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
