using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBProjects : ISafeComWrapper, IComCollection<IVBProject>, IEquatable<IVBProjects>
    {
        IVBE VBE { get; }
        IVBE Parent { get; }
        IVBProject Add(ProjectType type);
        IVBProject Open(string path);
        void Remove(IVBProject project);

        IVBProjectsEventsSink Events { get; }
        IConnectionPoint ConnectionPoint { get; }
    }
}