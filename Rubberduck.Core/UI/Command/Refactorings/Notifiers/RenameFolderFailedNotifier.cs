using Rubberduck.Interaction;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class RenameFolderFailedNotifier : RefactoringFailureNotifierBase
    {
        public RenameFolderFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => Resources.RubberduckUI.RenameDialog_Caption;
    }
}