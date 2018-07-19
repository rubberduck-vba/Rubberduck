using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComInterface : ComType, IComTypeWithMembers
    {
        public bool IsExtensible { get; private set; }

        private readonly List<ComInterface> _inherited = new List<ComInterface>();
        public IEnumerable<ComInterface> InheritedInterfaces => _inherited;

        private readonly List<ComMember> _members = new List<ComMember>();
        public IEnumerable<ComMember> Members => _members;

        private readonly List<ComField> _properties = new List<ComField>();
        public IEnumerable<ComField> Properties => _properties;

        public ComMember DefaultMember { get; private set; }

        public ComInterface(ITypeInfo info, TYPEATTR attrib) : base(info, attrib)
        {
            GetImplementedInterfaces(info, attrib);
            GetComProperties(info, attrib);
            GetComMembers(info, attrib);            
        }

        public ComInterface(ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(typeLib, attrib, index)
        {
            Type = DeclarationType.ClassModule;
            GetImplementedInterfaces(info, attrib);
            GetComProperties(info, attrib);
            GetComMembers(info, attrib);            
        }

        private void GetImplementedInterfaces(ITypeInfo info, TYPEATTR typeAttr)
        {
            IsExtensible = !typeAttr.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FNONEXTENSIBLE);
            for (var implIndex = 0; implIndex < typeAttr.cImplTypes; implIndex++)
            {
                info.GetRefTypeOfImplType(implIndex, out int href);
                info.GetRefTypeInfo(href, out ITypeInfo implemented);

                implemented.GetTypeAttr(out IntPtr attribPtr);
                using (DisposalActionContainer.Create(attribPtr, info.ReleaseTypeAttr))
                {
                    var attribs = Marshal.PtrToStructure<TYPEATTR>(attribPtr);

                    ComProject.KnownTypes.TryGetValue(attribs.guid, out ComType inherited);
                    var intface = inherited as ComInterface ?? new ComInterface(implemented, attribs);
                    _inherited.Add(intface);
                    ComProject.KnownTypes.TryAdd(attribs.guid, intface);
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
                    var comMember = new ComMember(info, member);
                    _members.Add(comMember);
                    if (comMember.IsDefault)
                    {
                        DefaultMember = comMember;
                    }
                }
            }
        }

        private void GetComProperties(ITypeInfo info, TYPEATTR attrib)
        {
            var names = new string[1];
            for (var index = 0; index < attrib.cVars; index++)
            {
                info.GetVarDesc(index, out IntPtr varDescPtr);
                using (DisposalActionContainer.Create(varDescPtr, info.ReleaseVarDesc))
                {
                    var property = Marshal.PtrToStructure<VARDESC>(varDescPtr);
                    info.GetNames(property.memid, names, names.Length, out int length);
                    Debug.Assert(length == 1);

                    _properties.Add(new ComField(info, property, names[0], index, DeclarationType.Property));
                }
            }
        }
    }
}
