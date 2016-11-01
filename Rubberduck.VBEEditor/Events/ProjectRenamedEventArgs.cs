using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ProjectRenamedEventArgs : ProjectEventArgs
    {
        public ProjectRenamedEventArgs(string projectId, IVBProject project, string oldName) : base(projectId, project)
        {
            _oldName = oldName;
        }

        private readonly string _oldName;
        public string OldName { get { return _oldName; } }
    }
}