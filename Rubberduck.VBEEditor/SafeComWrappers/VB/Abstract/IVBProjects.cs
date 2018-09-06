using System;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBProjects : ISafeEventedComWrapper, IComCollection<IVBProject>, IEquatable<IVBProjects>
    {
        event EventHandler<ProjectEventArgs> ProjectActivated;
        event EventHandler<ProjectEventArgs> ProjectAdded;
        event EventHandler<ProjectEventArgs> ProjectRemoved;
        event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;

        IVBE VBE { get; }
        IVBE Parent { get; }
        IVBProject Add(ProjectType type);
        IVBProject Open(string path);
        void Remove(IVBProject project);
        IVBProject StartProject { get; set; } // VB6 only
    }
}