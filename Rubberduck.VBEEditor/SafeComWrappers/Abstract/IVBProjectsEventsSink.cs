using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBProjectsEventsSink
    {
        event EventHandler<DispatcherEventArgs<IVBProject>> ProjectAdded;
        event EventHandler<DispatcherEventArgs<IVBProject>> ProjectRemoved;
        event EventHandler<DispatcherRenamedEventArgs<IVBProject>> ProjectRenamed;
        event EventHandler<DispatcherEventArgs<IVBProject>> ProjectActivated;
    }
}