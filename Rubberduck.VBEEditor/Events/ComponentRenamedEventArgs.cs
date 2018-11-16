using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ComponentRenamedEventArgs : ComponentEventArgs
    {
        public ComponentRenamedEventArgs(QualifiedModuleName qmn, string oldName)
            : base(qmn)
        {
            OldName = oldName;
        }

        public string OldName { get; }
    }
}