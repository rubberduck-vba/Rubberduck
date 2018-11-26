using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [KnownType(typeof(ComType))]
    public class ComStruct : ComType, IComTypeWithFields
    {
        [DataMember(IsRequired = true)]
        private List<ComField> _fields = new List<ComField>();
        public IEnumerable<ComField> Fields => _fields;

        public ComStruct(IComBase parent, ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index)
            : base(parent, typeLib, attrib, index)
        {
            Type = DeclarationType.UserDefinedType;          
            GetFields(info, attrib);
        }

        private void GetFields(ITypeInfo info, TYPEATTR attrib)
        {
            var names = new string[1];
            for (var index = 0; index < attrib.cVars; index++)
            {
                info.GetVarDesc(index, out IntPtr varPtr);
                using (DisposalActionContainer.Create(varPtr, info.ReleaseVarDesc))
                {
                    var desc = Marshal.PtrToStructure<VARDESC>(varPtr);
                    info.GetNames(desc.memid, names, names.Length, out int length);
                    Debug.Assert(length == 1);

                    _fields.Add(new ComField(this, info, names[0], desc, index, DeclarationType.UserDefinedTypeMember));
                }
            }
        }
    }
}
