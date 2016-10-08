using System;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.Parsing
{
    public interface IProjectEventArgs
    {
        string ProjectId { get; }
        VBProject Project { get; }
    }

    public interface IProjectRenamedEventArgs : IProjectEventArgs
    {
        string OldName { get; }
    }

    public interface IComponentEventArgs
    {
        string ProjectId { get; }

        VBProject Project { get; }
        VBComponent Component { get; }
    }

    public interface IComponentRenamedEventArgs : IComponentEventArgs
    {
        string OldName { get; }
    }

    public interface ISinks
    {
        void Start();
        bool ComponentSinksEnabled { get; set; }

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
