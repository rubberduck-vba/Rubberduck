using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ComponentEventArgs : EventArgs
    {
        public ComponentEventArgs(string projectId, IVBProject project, IVBComponent component)
        {
            ProjectId = projectId;
            Project = project;
            Component = component;
        }

        public string ProjectId { get; }

        public IVBProject Project { get; }

        public IVBComponent Component { get; }
    }
}