using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBProjects : ISafeComWrapper, IComCollection<IVBProject>, IEquatable<IVBProjects>
    {
        IVBE VBE { get; }
        IVBE Parent { get; }
        IVBProject Add(ProjectType type);
        IVBProject Open(string path);
        void Remove(Microsoft.Vbe.Interop.VBProject project);
    }
}