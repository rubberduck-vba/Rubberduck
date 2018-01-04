﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComInterface : ComType, IComTypeWithMembers
    {
        private readonly List<ComInterface> _inherited = new List<ComInterface>();
        private readonly List<ComMember> _members = new List<ComMember>();

        public bool IsExtensible { get; private set; }

        public IEnumerable<ComInterface> InheritedInterfaces => _inherited;

        public IEnumerable<ComMember> Members => _members;

        public ComMember DefaultMember { get; private set; }

        public ComInterface(ITypeInfo info, TYPEATTR attrib) : base(info, attrib)
        {
            GetImplementedInterfaces(info, attrib);
            GetComMembers(info, attrib);
        }

        public ComInterface(ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(typeLib, attrib, index)
        {
            Type = DeclarationType.ClassModule;
            GetImplementedInterfaces(info, attrib);
            GetComMembers(info, attrib);
        }

        private void GetImplementedInterfaces(ITypeInfo info, TYPEATTR typeAttr)
        {
            IsExtensible = !typeAttr.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FNONEXTENSIBLE);
            for (var implIndex = 0; implIndex < typeAttr.cImplTypes; implIndex++)
            {
                info.GetRefTypeOfImplType(implIndex, out var href);

                info.GetRefTypeInfo(href, out var implemented);

                implemented.GetTypeAttr(out var attribPtr);
                var attribs = (TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(TYPEATTR));

                ComProject.KnownTypes.TryGetValue(attribs.guid, out var inherited);
                var intface = inherited as ComInterface ?? new ComInterface(implemented, attribs);
                _inherited.Add(intface);
                ComProject.KnownTypes.TryAdd(attribs.guid, intface);

                info.ReleaseTypeAttr(attribPtr);
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

                var comMember = new ComMember(info, member);
                _members.Add(comMember);
                if (comMember.IsDefault)
                {
                    DefaultMember = comMember;
                }
                info.ReleaseFuncDesc(memberPtr);
            }
        }
    }
}
