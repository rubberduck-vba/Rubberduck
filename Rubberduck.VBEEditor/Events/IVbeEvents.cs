﻿using System;

namespace Rubberduck.VBEditor.Events
{
    public interface IVBEEvents
    {
        event EventHandler<ProjectEventArgs> ProjectAdded;
        event EventHandler<ProjectEventArgs> ProjectRemoved;
        event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;
        event EventHandler<ProjectEventArgs> ProjectActivated;
        event EventHandler<ComponentEventArgs> ComponentAdded;
        event EventHandler<ComponentEventArgs> ComponentRemoved;
        event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        event EventHandler<ComponentEventArgs> ComponentSelected;
        event EventHandler<ComponentEventArgs> ComponentActivated;
        event EventHandler<ComponentEventArgs> ComponentReloaded;
        event EventHandler EventsTerminated;
    }
}
