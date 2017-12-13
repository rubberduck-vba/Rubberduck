using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ProjectEventArgs : EventArgs
    {
        public ProjectEventArgs(string projectId, IVBProject project)
        {
            ProjectId = projectId;
            Project = project;
        }

        public string ProjectId { get; }

        public IVBProject Project { get; }
    }
}