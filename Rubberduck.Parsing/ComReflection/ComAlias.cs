using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name} As {TypeName}")]
    public class ComAlias : ComBase
    {
        public VarEnum VarType { get; private set; }
        public string TypeName { get; private set; }
        public bool IsHidden { get; private set; }
        public bool IsRestricted { get; private set; }

        public ComAlias(ITypeLib typeLib, ITypeInfo info, int index, TYPEATTR attributes) : base(typeLib, index)
        {
            Index = index;
            Documentation = new ComDocumentation(typeLib, index);
            Guid = attributes.guid;
            VarType = (VarEnum)attributes.tdescAlias.vt;
            IsHidden = attributes.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN);
            IsRestricted = attributes.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FRESTRICTED);
            
            if (Name.Equals("LONG_PTR"))
            {
                TypeName = "LongPtr";
                return;                
            }
          
            if (ComVariant.TypeNames.ContainsKey(VarType))
            {
                TypeName = ComVariant.TypeNames[VarType];
            }
            else if (VarType == VarEnum.VT_USERDEFINED)
            {
                ITypeInfo refType;
                info.GetRefTypeInfo((int)attributes.tdescAlias.lpValue, out refType);
                var doc = new ComDocumentation(refType, -1);
                TypeName = doc.Name;
            }
            else
            {
                throw new NotImplementedException(string.Format("Didn't expect an alias with a type of {0}.", VarType));
            }
        }
    }
}
