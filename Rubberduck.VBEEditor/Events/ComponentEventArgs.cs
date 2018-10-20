using System;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ComponentEventArgs : EventArgs
    {
        private readonly QualifiedModuleName _qmn;

        public ComponentEventArgs(QualifiedModuleName qmn)
        {
            _qmn = qmn;
        }

        public string ProjectId => _qmn.ProjectId;
        public QualifiedModuleName QualifiedModuleName => _qmn;

        public bool TryGetProject(IProjectsRepository repository, out IVBProject project)
        {
            try
            {
                project = repository.Project(_qmn.ProjectId);
                return true;
            }
            catch
            {
                project = null;
                return false;
            }
        }

        public bool TryGetComponent(IProjectsRepository repository, out IVBComponent component)
        {
            try
            {
                component = repository.Component(_qmn);
                return true;
            }
            catch
            {
                component = null;
                return false;
            }
        }
    }
}