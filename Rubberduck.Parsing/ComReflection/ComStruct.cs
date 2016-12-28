using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComStruct : ComType, IComTypeWithFields
    {
        private readonly List<ComField> _fields = new List<ComField>();
        public IEnumerable<ComField> Fields
        {
            get { return _fields; }
        }

        public ComStruct(ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index)
            : base(typeLib, attrib, index)
        {
            _fields = new List<ComField>();
            Type = DeclarationType.UserDefinedType;          
            GetFields(info, attrib);
        }

        private void GetFields(ITypeInfo info, TYPEATTR attrib)
        {
            var names = new string[255];
            for (var index = 0; index < attrib.cVars; index++)
            {
                IntPtr varPtr;
                info.GetVarDesc(index, out varPtr);
                var desc = (VARDESC)Marshal.PtrToStructure(varPtr, typeof(VARDESC));
                int length;
                info.GetNames(desc.memid, names, 255, out length);
                Debug.Assert(length == 1);

                _fields.Add(new ComField(names[0], desc, index, DeclarationType.UserDefinedTypeMember));
                info.ReleaseVarDesc(varPtr);
            }
        }
    }
}
