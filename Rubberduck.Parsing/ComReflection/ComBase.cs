using System;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
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
    }

    public abstract class ComBase : IComBase
    {
        public Guid Guid { get; protected set; }
        public int Index { get; protected set; }
        public ComDocumentation Documentation { get; protected set; }
        public string Name
        {
            get { return Documentation == null ? string.Empty : Documentation.Name; }
        }

        public DeclarationType Type { get; protected set; }

        protected ComBase(ITypeLib typeLib, int index)
        {
            Index = index;
            Documentation = new ComDocumentation(typeLib, index);
        }

        protected ComBase(ITypeInfo info)
        {
            ITypeLib typeLib;
            int index;
            info.GetContainingTypeLib(out typeLib, out index);
            Index = index;
            Debug.Assert(typeLib != null);
            Documentation = new ComDocumentation(typeLib, index);
        }

        protected ComBase(ITypeInfo info, FUNCDESC funcDesc)
        {
            Index = funcDesc.memid;
            Documentation = new ComDocumentation(info, funcDesc.memid);
        }
    }
}
