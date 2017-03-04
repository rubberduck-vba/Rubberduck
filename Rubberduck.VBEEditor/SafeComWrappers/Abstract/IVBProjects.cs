using System;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBProjects : ISafeComWrapper, IComCollection<IVBProject>, IEquatable<IVBProjects>
    {
        //event EventHandler<ProjectEventArgs> ProjectActivated;
        //event EventHandler<ProjectEventArgs> ProjectAdded;
        //event EventHandler<ProjectEventArgs> ProjectRemoved;
        //event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;

        IVBE VBE { get; }
        IVBE Parent { get; }
        IVBProject Add(ProjectType type);
        IVBProject Open(string path);
        void Remove(IVBProject project);
    }
}