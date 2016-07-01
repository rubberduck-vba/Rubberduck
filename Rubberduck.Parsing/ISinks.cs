using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public interface IProjectEventArgs
    {
        string ProjectId { get; }
    }

    public interface IProjectRenamedEventArgs : IProjectEventArgs
    {
        string OldName { get; }
    }

    public interface IComponentEventArgs
    {
        string ProjectId { get; }
        string ComponentName { get; }
        vbext_ComponentType Type { get; }
    }

    public interface IComponentRenamedEventArgs : IComponentEventArgs
    {
        string OldName { get; }
    }

    public interface ISinks
    {
        bool IsEnabled { get; set; }

        event EventHandler<IProjectEventArgs> ProjectActivated;
        event EventHandler<IProjectEventArgs> ProjectAdded;
        event EventHandler<IProjectEventArgs> ProjectRemoved;
        event EventHandler<IProjectRenamedEventArgs> ProjectRenamed;

        event EventHandler<IComponentEventArgs> ComponentActivated;
        event EventHandler<IComponentEventArgs> ComponentAdded;
        event EventHandler<IComponentEventArgs> ComponentReloaded;
        event EventHandler<IComponentEventArgs> ComponentRemoved;
        event EventHandler<IComponentRenamedEventArgs> ComponentRenamed;
        event EventHandler<IComponentEventArgs> ComponentSelected;

        //event EventHandler<IDispatcherEventArgs<Reference>> ReferenceAdded;
        //event EventHandler<IDispatcherEventArgs<Reference>> ReferenceRemoved;
    }
}
