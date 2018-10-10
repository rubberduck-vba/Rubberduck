using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComEnumeration : ComType
    {
        public List<ComEnumerationMember> Members { get; }

        public ComEnumeration(IComBase parent, ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(parent, typeLib, attrib, index)
        {
            Members = new List<ComEnumerationMember>();
            Type = DeclarationType.Enumeration;
            GetEnumerationMembers(info, attrib);
            ComProject.KnownEnumerations.TryAdd(Guid, this);
        }

        private void GetEnumerationMembers(ITypeInfo info, TYPEATTR attrib)
        {
            var count = attrib.cVars;
            for (var index = 0; index < count; index++)
            {
                info.GetVarDesc(index, out IntPtr varPtr);
                using (DisposalActionContainer.Create(varPtr, info.ReleaseVarDesc))
                {
                    var desc = Marshal.PtrToStructure<VARDESC>(varPtr);
                    Members.Add(new ComEnumerationMember(this, info, desc));
                }
            }
        }
    }
}
