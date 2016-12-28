using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComEnumeration : ComType
    {
        public List<ComEnumerationMember> Members { get; set; } 

        public ComEnumeration(ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(typeLib, attrib, index)
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
                IntPtr varPtr;
                info.GetVarDesc(index, out varPtr);
                var desc = (VARDESC)Marshal.PtrToStructure(varPtr, typeof(VARDESC));
                Members.Add(new ComEnumerationMember(info, desc));
                info.ReleaseVarDesc(varPtr);
            }           
        }
    }
}
