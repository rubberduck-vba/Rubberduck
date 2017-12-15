using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ProjectRenamedEventArgs : ProjectEventArgs
    {
        public ProjectRenamedEventArgs(string projectId, IVBProject project, string oldName) : base(projectId, project)
        {
            OldName = oldName;
        }

        public string OldName { get; }
    }
}