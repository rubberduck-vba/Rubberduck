using System;

namespace Rubberduck.VBEditor.Events
{
    public interface ISinks
    {
        void Start();
        void Stop();
        bool ComponentSinksEnabled { get; set; }

        event EventHandler<ProjectEventArgs> ProjectActivated;
        event EventHandler<ProjectEventArgs> ProjectAdded;
        event EventHandler<ProjectEventArgs> ProjectRemoved;
        event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;

        event EventHandler<ComponentEventArgs> ComponentActivated;
        event EventHandler<ComponentEventArgs> ComponentAdded;
        event EventHandler<ComponentEventArgs> ComponentReloaded;
        event EventHandler<ComponentEventArgs> ComponentRemoved;
        event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        event EventHandler<ComponentEventArgs> ComponentSelected;
    }
}
