using System;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Symbols;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComBase
    {
        Guid Guid { get; }
        int Index { get; }
        ComDocumentation Documentation { get; }
        string Name { get; }
        DeclarationType Type { get; }
        IComBase  Parent { get; }
        ComProject Project { get; }
    }

    [DataContract]
    public abstract class ComBase : IComBase
    {
        [DataMember(IsRequired = true)]
        public Guid Guid { get; protected set; }

        [DataMember(IsRequired = true)]
        public int Index { get; protected set; }

        [DataMember(IsRequired = true)]
        public ComDocumentation Documentation { get; protected set; }

        public string Name => Documentation == null ? string.Empty : Documentation.Name;

        [DataMember(IsRequired = true)]
        public DeclarationType Type { get; protected set; }

        [DataMember(IsRequired = true)]
        public IComBase Parent { get; protected set; }

        public ComProject Project => Parent != null ? Parent.Project : this as ComProject;

        protected ComBase(IComBase parent, ITypeLib typeLib, int index)
        {
            Parent = parent;
            Index = index;
            Documentation = new ComDocumentation(typeLib, index);
        }

        protected ComBase(IComBase parent, ITypeInfo info)
        {
            Parent = parent;
            info.GetContainingTypeLib(out ITypeLib typeLib, out int index);
            Index = index;
            Debug.Assert(typeLib != null);
            Documentation = new ComDocumentation(typeLib, index);
        }

        protected ComBase(IComBase parent, ITypeInfo info, FUNCDESC funcDesc)
        {
            Parent = parent;
            Index = funcDesc.memid;
            Documentation = new ComDocumentation(info, funcDesc.memid);
        }
    }
}
