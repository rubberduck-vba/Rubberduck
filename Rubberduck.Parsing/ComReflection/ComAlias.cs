using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComAlias : ComBase
    {
        public VarEnum VarType { get; set; }

        public string TypeName { get; set; }

        public ComAlias(ITypeLib typeLib, ITypeInfo info, int index, TYPEATTR attributes) : base(typeLib, index)
        {
            Index = index;
            Documentation = new ComDocumentation(typeLib, index);
            VarType = (VarEnum)attributes.tdescAlias.vt;
            if (ComVariant.TypeNames.ContainsKey(VarType))
            {
                TypeName = ComVariant.TypeNames[VarType];
            }
            else if (VarType == VarEnum.VT_USERDEFINED)
            {
                //?
            }
            else
            {
                throw new NotImplementedException(string.Format("Didn't expect an alias with a type of {0}.", VarType));
            }
        }
    }
}
