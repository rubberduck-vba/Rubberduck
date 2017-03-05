using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComType : IComBase
    {
        bool IsAppObject { get; }
        bool IsPreDeclared { get; }
        bool IsHidden { get; }
        bool IsRestricted { get; }
    }

    public interface IComTypeWithMembers : IComType
    {
        IEnumerable<ComMember> Members { get; }
        ComMember DefaultMember { get; }
        bool IsExtensible { get; }
    }

    public interface IComTypeWithFields : IComType
    {
        IEnumerable<ComField> Fields { get; }
    }

    [DebuggerDisplay("{Name}")]
    public abstract class ComType : ComBase, IComType
    {
        public bool IsAppObject { get; private set; }
        public bool IsPreDeclared { get; private set; }
        public bool IsHidden { get; private set; }
        public bool IsRestricted { get; private set; }
        
        protected ComType(ITypeInfo info, TYPEATTR attrib)
            : base(info)
        {
            SetFlagsFromTypeAttr(attrib);
        }

        protected ComType(ITypeLib typeLib, TYPEATTR attrib, int index)
            : base(typeLib, index)
        {
            Index = index;
            SetFlagsFromTypeAttr(attrib);
        }

        private void SetFlagsFromTypeAttr(TYPEATTR attrib)
        {
            Guid = attrib.guid;
            IsAppObject = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FAPPOBJECT);
            IsPreDeclared = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FPREDECLID);
            IsHidden = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN);
            IsRestricted = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FRESTRICTED);           
        }
    }
}
