using Rubberduck.Interaction;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class ExtractMethodFailedNotifier : RefactoringFailureNotifierBase
    {
        public ExtractMethodFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => RefactoringsUI.ExtractMethod_Caption;
    }
}