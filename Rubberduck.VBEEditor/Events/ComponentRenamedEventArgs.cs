using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ComponentRenamedEventArgs : ComponentEventArgs
    {
        public ComponentRenamedEventArgs(string projectId, IVBProject project, IVBComponent component, string oldName)
            : base(projectId, project, component)
        {
            _oldName = oldName;
        }

        private readonly string _oldName;
        public string OldName { get { return _oldName; } }
    }
}