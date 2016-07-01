using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public interface IDispatcherEventArgs<T>
        where T : class
    {
        T Item { get; }
    }

    public interface IDispatcherRenamedEventArgs<T> : IDispatcherEventArgs<T>
        where T : class
    {
        string OldName { get; }
    }

    public interface ISinks
    {
        bool IsEnabled { get; set; }

        event EventHandler<IDispatcherEventArgs<VBProject>> ProjectActivated;
        event EventHandler<IDispatcherEventArgs<VBProject>> ProjectAdded;
        event EventHandler<IDispatcherEventArgs<VBProject>> ProjectRemoved;
        event EventHandler<IDispatcherRenamedEventArgs<VBProject>> ProjectRenamed;

        event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentActivated;
        event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentAdded;
        event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentReloaded;
        event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentRemoved;
        event EventHandler<IDispatcherRenamedEventArgs<VBComponent>> ComponentRenamed;
        event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentSelected;

        //event EventHandler<IDispatcherEventArgs<Reference>> ReferenceAdded;
        //event EventHandler<IDispatcherEventArgs<Reference>> ReferenceRemoved;
    }
}
