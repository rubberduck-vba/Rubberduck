using Rubberduck.Interaction;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class RenameFolderFailedNotifier : RefactoringFailureNotifierBase
    {
        public RenameFolderFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => RefactoringsUI.RenameDialog_Caption;
    }
}