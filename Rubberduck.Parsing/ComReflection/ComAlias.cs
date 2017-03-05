using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name} As {TypeName}")]
    public class ComAlias : ComBase
    {
        public string TypeName { get; private set; }
        public bool IsHidden { get; private set; }
        public bool IsRestricted { get; private set; }

        public ComAlias(ITypeLib typeLib, ITypeInfo info, int index, TYPEATTR attributes) : base(typeLib, index)
        {
            Index = index;
            Documentation = new ComDocumentation(typeLib, index);
            Guid = attributes.guid;
            IsHidden = attributes.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN);
            IsRestricted = attributes.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FRESTRICTED);
            
            if (Name.Equals("LONG_PTR"))
            {
                TypeName = "LongPtr";
                return;                
            }

            var aliased = new ComParameter(attributes, info);
            TypeName = aliased.TypeName;
        }
    }
}
