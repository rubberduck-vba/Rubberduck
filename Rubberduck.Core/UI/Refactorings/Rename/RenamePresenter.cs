using System;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.Rename
{
    internal class RenamePresenter : RefactoringPresenterBase<RenameModel>, IRenamePresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.RenameDialog_Caption, 164, 684);

        private readonly IMessageBox _messageBox;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RenamePresenter(RenameModel model, IRefactoringDialogFactory dialogFactory, IMessageBox messageBox) :
            base(DialogData, model, dialogFactory)
        {
            _messageBox = messageBox;
        }

        public override RenameModel Show()
        {
            if (Model?.Target == null)
            {
                return null;
            }

            if (!Model.Target.Equals(Model.InitialTarget)
                && !UserConfirmsNewTarget(Model))
            {
                throw new RefactoringAbortedException();
            }

            return base.Show();
        }

        public RenameModel Show(Declaration target)
        {
            if (null == target)
            {
                return null;
            }

            Model.Target = target;

            return Show();
        }

        private bool UserConfirmsNewTarget(RenameModel model)
        {
            var initialTarget = model.InitialTarget;
            var newTarget = model.Target;

            if (model.IsControlEventHandlerRename)
            {
                var message = string.Format(RubberduckUI.RenamePresenter_TargetIsControlEventHandler, initialTarget.IdentifierName, newTarget.IdentifierName);
                return UserConfirmsRenameOfResolvedTarget(message);
            }

            if (model.IsUserEventHandlerRename)
            {
                var message = string.Format(RubberduckUI.RenamePresenter_TargetIsEventHandlerImplementation, initialTarget.IdentifierName, newTarget.ComponentName, newTarget.IdentifierName);
                return UserConfirmsRenameOfResolvedTarget(message);
            }

            if (model.IsInterfaceMemberRename)
            {
                var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, initialTarget.IdentifierName, newTarget.ComponentName, newTarget.IdentifierName);
                return UserConfirmsRenameOfResolvedTarget(message);
            }

            _logger.Error("Unexpected resolution to different target declaration in RenameRefactoring.");
            _logger.Debug($"original target: {initialTarget.QualifiedName}{Environment.NewLine}new target: {newTarget.QualifiedName}");
            return false;
        }

        private bool UserConfirmsRenameOfResolvedTarget(string message)
        {
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_TitleText) ?? false;
        }
    }
}

