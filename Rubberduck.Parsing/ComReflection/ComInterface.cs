using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [KnownType(typeof(ComType))]
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComInterface : ComType, IComTypeWithMembers
    {
        [DataMember(IsRequired = true)]
        public bool IsExtensible { get; private set; }

        [DataMember(IsRequired = true)]
        private List<ComInterface> _inherited = new List<ComInterface>();
        public IEnumerable<ComInterface> InheritedInterfaces => _inherited;

        [DataMember(IsRequired = true)]
        private List<ComMember> _members = new List<ComMember>();
        public IEnumerable<ComMember> Members => _members;

        [DataMember(IsRequired = true)]
        private List<ComField> _properties = new List<ComField>();
        public IEnumerable<ComField> Properties => _properties;

        [DataMember(IsRequired = true)]
        public ComMember DefaultMember { get; private set; }

        public ComInterface(IComBase parent, ITypeInfo info, TYPEATTR attrib) : base(parent, info, attrib)
        {
            // Since the reference declaration gathering is threaded, this can't be truly recursive, so implemented interfaces may have
            // null parents (for example, if the library references a type library that isn't referenced by the VBA project or if that project
            // hasn't had the implemented interface processed yet).
            try
            {
                info.GetContainingTypeLib(out var typeLib, out _);
                typeLib.GetLibAttr(out IntPtr attribPtr);
                using (DisposalActionContainer.Create(attribPtr, typeLib.ReleaseTLibAttr))
                {
                    var typeAttr = Marshal.PtrToStructure<TYPELIBATTR>(attribPtr);
                    Parent = typeAttr.guid.Equals(parent?.Project.Guid) ? parent?.Project : null;
                }
            }
            catch
            {
                Parent = null;
            }

            GetImplementedInterfaces(info, attrib);
            GetComProperties(info, attrib);
            GetComMembers(info, attrib);
        }

        public ComInterface(IComBase parent, ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base(parent, typeLib, attrib, index)
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
                using (DisposalActionContainer.Create(attribPtr, implemented.ReleaseTypeAttr))
                {
                    var attribs = Marshal.PtrToStructure<TYPEATTR>(attribPtr);

                    ComProject.KnownTypes.TryGetValue(attribs.guid, out ComType inherited);
                    var intface = inherited as ComInterface ?? new ComInterface(Project, implemented, attribs);
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
                    var comMember = new ComMember(this, info, member);
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

                    _properties.Add(new ComField(this, info, names[0], property, index, DeclarationType.Property));
                }
            }
        }
    }
}
