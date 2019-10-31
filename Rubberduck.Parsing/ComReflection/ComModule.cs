using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [KnownType(typeof(ComType))]
    public class ComModule : ComType, IComTypeWithMembers, IComTypeWithFields
    {
        [DataMember(IsRequired = true)]
        private List<ComMember> _members = new List<ComMember>();
        public IEnumerable<ComMember> Members => _members;

        public ComMember DefaultMember => null;

        public bool IsExtensible => false;

        [DataMember(IsRequired = true)]
        private List<ComField> _fields = new List<ComField>();
        public IEnumerable<ComField> Fields => _fields;

        public IEnumerable<ComField> Properties => Enumerable.Empty<ComField>();

        public ComModule(IComBase parent, ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(parent, typeLib, attrib, index)
        {
            Debug.Assert(attrib.cFuncs >= 0 && attrib.cVars >= 0);
            Type = DeclarationType.ProceduralModule;
            if (attrib.cFuncs > 0)
            {
                GetComMembers(info, attrib);
            }
            if (attrib.cVars > 0)
            {
                GetComFields(info, attrib);
            }
        }

        private void GetComFields(ITypeInfo info, TYPEATTR attrib)
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

                    DeclarationType type;
                    if(info is ITypeInfoWrapper wrapped && wrapped.HasVBEExtensions)
                    {
                        type = desc.IsValidVBAConstant() ? DeclarationType.Constant : DeclarationType.Variable;
                    }
                    else
                    {
                        type = desc.varkind == VARKIND.VAR_CONST ? DeclarationType.Constant : DeclarationType.Variable;
                    }

                    _fields.Add(new ComField(this, info, names[0], desc, index, type));
                }
            }
        }

        private void GetComMembers(ITypeInfo info, TYPEATTR attrib)
        {
            for (var index = 0; index < attrib.cFuncs; index++)
            {
                info.GetFuncDesc(index, out IntPtr memberPtr);
                using (DisposalActionContainer.Create(memberPtr, info.ReleaseFuncDesc))
                {
                    var member = Marshal.PtrToStructure<FUNCDESC>(memberPtr);
                    if (member.callconv != CALLCONV.CC_STDCALL)
                    {
                        continue;
                    }
                    _members.Add(new ComMember(this, info, member));
                }
            }
        }
    }
}
