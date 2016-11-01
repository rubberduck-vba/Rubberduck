using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ComponentEventArgs : EventArgs
    {
        public ComponentEventArgs(string projectId, IVBProject project, IVBComponent component)
        {
            _projectId = projectId;
            _project = project;
            _component = component;
        }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly IVBProject _project;
        public IVBProject Project { get { return _project; } }

        private readonly IVBComponent _component;
        public IVBComponent Component { get { return _component; } }

    }
}