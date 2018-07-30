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
        IEnumerable<ComField> Properties { get; }
        ComMember DefaultMember { get; }
        bool IsExtensible { get; }
    }

    public interface IComTypeWithFields : IComType
    {
        IEnumerable<ComField> Fields { get; }
    }

    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public abstract class ComType : ComBase, IComType
    {
        public bool IsAppObject { get; }
        public bool IsPreDeclared { get; }
        public bool IsHidden { get; }
        public bool IsRestricted { get; }
        
        protected ComType(IComBase parent, ITypeInfo info, TYPEATTR attrib)
            : base(parent, info)
        {
            Guid = attrib.guid;
            IsAppObject = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FAPPOBJECT);
            IsPreDeclared = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FPREDECLID);
            IsHidden = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN);
            IsRestricted = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FRESTRICTED);
        }

        protected ComType(IComBase parent, ITypeLib typeLib, TYPEATTR attrib, int index)
            : base(parent, typeLib, index)
        {
            Index = index;
            Guid = attrib.guid;
            IsAppObject = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FAPPOBJECT);
            IsPreDeclared = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FPREDECLID);
            IsHidden = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN);
            IsRestricted = attrib.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FRESTRICTED);
        }
    }
}
