﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComModule : ComType, IComTypeWithMembers, IComTypeWithFields
    {
        private readonly List<ComMember> _members = new List<ComMember>();
        public IEnumerable<ComMember> Members => _members;

        public ComMember DefaultMember => null;

        public bool IsExtensible => false;

        private readonly List<ComField> _fields = new List<ComField>();
        public IEnumerable<ComField> Fields => _fields;

        public ComModule(ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(typeLib, attrib, index)
        {
            Type = DeclarationType.ProceduralModule;
            if (attrib.cFuncs > 0)
            {
                Debug.Assert(attrib.cVars == 0);
                GetComMembers(info, attrib);
            }
            else
            {
                Debug.Assert(attrib.cVars > 0);
                GetComFields(info, attrib);
            }
        }

        private void GetComFields(ITypeInfo info, TYPEATTR attrib)
        {
            var names = new string[255];
            for (var index = 0; index < attrib.cVars; index++)
            {
                info.GetVarDesc(index, out var varPtr);
                var desc = (VARDESC)Marshal.PtrToStructure(varPtr, typeof(VARDESC));
                info.GetNames(desc.memid, names, 255, out var length);
                Debug.Assert(length == 1);

                _fields.Add(new ComField(names[0], desc, index, DeclarationType.Constant));
                info.ReleaseVarDesc(varPtr);
            }
        }

        private void GetComMembers(ITypeInfo info, TYPEATTR attrib)
        {
            for (var index = 0; index < attrib.cFuncs; index++)
            {
                info.GetFuncDesc(index, out var memberPtr);
                var member = (FUNCDESC)Marshal.PtrToStructure(memberPtr, typeof(FUNCDESC));
                if (member.callconv != CALLCONV.CC_STDCALL)
                {
                    continue;
                }
                _members.Add(new ComMember(info, member));
                info.ReleaseFuncDesc(memberPtr);
            }
        }
    }
}
