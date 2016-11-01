using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBComponentsEventsSink
    {
        event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentAdded;
        event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentRemoved;
        event EventHandler<DispatcherRenamedEventArgs<IVBComponent>> ComponentRenamed;
        event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentSelected;
        event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentActivated;
        event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentReloaded;
    }
}