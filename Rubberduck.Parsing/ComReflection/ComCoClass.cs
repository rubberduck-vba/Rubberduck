using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [KnownType(typeof(ComType))]
    public class ComCoClass : ComType, IComTypeWithMembers
    {
        [DataMember(IsRequired = true)]
        private Dictionary<ComInterface, bool> _interfaces = new Dictionary<ComInterface, bool>();

        [DataMember(IsRequired = true)]
        private List<ComInterface> _events = new List<ComInterface>();

        [DataMember(IsRequired = true)]
        public bool IsControl { get; private set; }

        public bool IsExtensible => _interfaces.Keys.Any(i => i.IsExtensible);

        [DataMember(IsRequired = true)]
        public ComInterface DefaultInterface { get; private set; }

        public IEnumerable<ComInterface> EventInterfaces => _events;

        public IEnumerable<ComInterface> ImplementedInterfaces => _interfaces.Keys;

        public IEnumerable<ComField> Properties => ImplementedInterfaces.Where(x => !_events.Contains(x)).SelectMany(i => i.Properties);

        public IEnumerable<ComInterface> VisibleInterfaces => _interfaces.Where(i => !i.Value).Select(i => i.Key);

        public IEnumerable<ComMember> Members => ImplementedInterfaces.Where(x => !_events.Contains(x)).SelectMany(i => i.Members);

        public ComMember DefaultMember => DefaultInterface.DefaultMember;

        public IEnumerable<ComMember> SourceMembers => _events.SelectMany(i => i.Members);

        public bool WithEvents => _events.Count > 0;

        public void AddInterface(ComInterface intrface, bool restricted = false)
        {
            Debug.Assert(intrface != null);
            if (!_interfaces.ContainsKey(intrface))
            {
                _interfaces.Add(intrface, restricted);
            }
        }

        public ComCoClass(IComBase parent, ITypeLib typeLib, ITypeInfo info, TYPEATTR attrib, int index) : base (parent, typeLib, attrib, index)
        {
            Type = DeclarationType.ClassModule;
            GetImplementedInterfaces(info, attrib);
            IsControl = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FCONTROL);
            Debug.Assert(attrib.cFuncs == 0);
        }

        private void GetImplementedInterfaces(ITypeInfo info, TYPEATTR typeAttr)
        {
            for (var implIndex = 0; implIndex < typeAttr.cImplTypes; implIndex++)
            {
                info.GetRefTypeOfImplType(implIndex, out int href);
                info.GetRefTypeInfo(href, out ITypeInfo implemented);

                implemented.GetTypeAttr(out IntPtr attribPtr);
                using (DisposalActionContainer.Create(attribPtr, info.ReleaseTypeAttr))
                {
                    var attribs = Marshal.PtrToStructure<TYPEATTR>(attribPtr);

                    ComProject.KnownTypes.TryGetValue(attribs.guid, out ComType inherited);
                    var intface = inherited as ComInterface ?? new ComInterface(Project, implemented, attribs);

                    ComProject.KnownTypes.TryAdd(attribs.guid, intface);

                    IMPLTYPEFLAGS flags = 0;
                    try
                    {
                        info.GetImplTypeFlags(implIndex, out flags);
                    }
                    catch (COMException) { }

                    if (flags.HasFlag(IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE))
                    {
                        _events.Add(intface);
                    }
                    else
                    {
                        DefaultInterface = flags.HasFlag(IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) ? intface : DefaultInterface;
                    }
                    _interfaces.Add(intface, flags.HasFlag(IMPLTYPEFLAGS.IMPLTYPEFLAG_FRESTRICTED));
                }
            }

            if (DefaultInterface == null)
            {
                DefaultInterface = VisibleInterfaces.FirstOrDefault();
            }
        }
    }
}
