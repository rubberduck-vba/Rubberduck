using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ProjectEventArgs : EventArgs
    {
        public ProjectEventArgs(string projectId, IVBProject project)
        {
            _projectId = projectId;
            _project = project;
        }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly IVBProject _project;
        public IVBProject Project { get { return _project; } }
    }
}