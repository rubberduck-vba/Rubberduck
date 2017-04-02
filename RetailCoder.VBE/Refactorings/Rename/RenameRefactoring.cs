using System.Linq;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Diagnostics;
using Microsoft.CSharp.RuntimeBinder;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private RenameModel _model;
        private bool _renameOperationIsCancelled;

        public RenameRefactoring(IVBE vbe, IRefactoringPresenterFactory<IRenamePresenter> factory, IMessageBox messageBox, RubberduckParserState state)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
            _state = state;
            _renameOperationIsCancelled = false;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();

            bool hasNullReferences = false;
            SetModelMember(presenter, out hasNullReferences);

            if(hasNullReferences) { return; }

            _model.Target = PreprocessSelectedTarget(_model.Target);

            if(_renameOperationIsCancelled) { return; }

            _model = presenter.Show(_model.Target);

            if (UserCancelsDueToNameConflict()) { return; }

            QualifiedSelection? qOldSelection = null;
            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                qOldSelection = pane.CodeModule.GetQualifiedSelection();
            }

            Rename();

            if (qOldSelection.HasValue)
            {
                pane.Selection = qOldSelection.Value.Selection;
            }
        }

        public void Refactor(QualifiedSelection qSelection)
        {
            var pane = _vbe.ActiveCodePane;
            if (pane.IsWrappingNullReference)
            {
                return;
            }
            pane.Selection = qSelection.Selection;
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            var presenter = _factory.Create();

            bool hasNullReferences = false;
            SetModelMember(presenter, out hasNullReferences);

            if (hasNullReferences) { return; }

            if (null == target)
            {
                if (null == _messageBox) { return; }
                _messageBox.Show(RubberduckUI.RefactorRename_TargetNotDefinedError, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            if (!target.IsUserDefined)
            {
                var message = string.Format(RubberduckUI.RefactorRename_TargetNotUserDefinedError, target.QualifiedName);
                _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            _model.Target = PreprocessSelectedTarget(_model.Target);

            if (_renameOperationIsCancelled) { return; }

            _model = presenter.Show(_model.Target);

            if (UserCancelsDueToNameConflict()) { return; }

            var uqOldSelection = Selection.Home;
            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                uqOldSelection = pane.Selection;
            }

            Rename();

            if (!pane.IsWrappingNullReference)
            {
                pane.Selection = uqOldSelection;
            }
        }

        private void SetModelMember(IRenamePresenter presenter, out bool hasNullReferences)
        {
            hasNullReferences = false;
            presenter = _factory.Create();
            _model = presenter.Model;

            if (null == _model)
            {
                hasNullReferences = true;
                return;
            }

            if (null == _model.Target)
            {
                hasNullReferences = true;
                if (null == _messageBox)
                {
                    return;
                }
                _messageBox.Show(RubberduckUI.RefactorRename_TargetNotDefinedError, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        //For Controls, Events, and Interfaces - make sure what the user is presented with
        //the declaration rather than a handler or implementation so that there is no confusion about what is 
        //being changed (e.g., If the user selects a control event handler like bnt1_Click, 
        //he is presented with 'bnt1' to rename.  Otherwise, if presented the handler, it is possible to 
        //make a change from 'bnt1_Click' to 'bnt1_ClickAgain' which would 
        //result in 'bnt1_ClickAgain_Click') 
        private Declaration PreprocessSelectedTarget(Declaration selectedTarget)
        {
            Declaration renameTarget = selectedTarget;

            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Procedure))
            {
                renameTarget = _model.ResolveHandlerToDeclaration(renameTarget, DeclarationType.Control);
                if( renameTarget.DeclarationType.HasFlag(DeclarationType.Control))
                {
                    return renameTarget;
                }
                Debug.Assert(renameTarget.DeclarationType.HasFlag(DeclarationType.Procedure), "Control Resolver Error: modified target type to " + renameTarget.DeclarationType.ToString());

                renameTarget = _model.ResolveHandlerToDeclaration(renameTarget, DeclarationType.Event);
                if (renameTarget.DeclarationType.HasFlag(DeclarationType.Event))
                {
                    return renameTarget;
                }
                Debug.Assert(renameTarget.DeclarationType.HasFlag(DeclarationType.Procedure), "Event Resolver Error: modified target type to " + renameTarget.DeclarationType.ToString());

                renameTarget = _model.ResolveImplementationToInterfaceDeclaration(renameTarget);

                if (selectedTarget != renameTarget)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, selectedTarget.IdentifierName, renameTarget.ComponentName, renameTarget.IdentifierName);

                    var confirm = _messageBox.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (confirm == DialogResult.No)
                    {
                        _renameOperationIsCancelled = true;
                    }
                }
                Debug.Assert(renameTarget.DeclarationType.HasFlag(DeclarationType.Procedure), "Interface Resolver Error: modified target type to " + renameTarget.DeclarationType.ToString());
            }

            return renameTarget;
        }

        private bool UserCancelsDueToNameConflict()
        {
            var declarations = _state.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(_model.Target)
                .Where(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase));//.FirstOrDefault();

            if (declarations.Any())
            {
                var rename = ConfirmProceedWithNameConflict(declarations.FirstOrDefault());

                if (rename == DialogResult.No)
                {
                    _renameOperationIsCancelled = true;
                }
            }
            return _renameOperationIsCancelled;
        }

        private DialogResult ConfirmProceedWithNameConflict(Declaration declaration)
        {
            var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                declaration);
            return _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation);
        }

        private void Rename()
        {
            var handler = GetHandler(_model, _messageBox);
            try
            {
                handler.Rename();
            }
            catch (RuntimeBinderException)
            {
                _messageBox.Show(handler.ErrorMessage, RubberduckUI.RenameDialog_Caption);
            }
            catch (COMException)
            {
                _messageBox.Show(handler.ErrorMessage, RubberduckUI.RenameDialog_Caption);
            }
        }

        private static IRenameRefactoringHandler GetHandler(RenameModel model, IMessageBox messageBox)
        {
            IRenameRefactoringHandler handler = null;
            if (model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenamePropertyHandler(model, messageBox);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Parameter)
                        && model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenamePropertyParameterHandler(model, messageBox);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                handler = new RenameModuleHandler(model, messageBox);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Project))
            {
                handler = new RenameProjectHandler(model, messageBox);
            }
            else if (model.Declarations.FindInterfaceMembers().Contains(model.Target))
            {
                handler = new RenameInterfaceHandler(model, messageBox);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Event)) 
            {
                handler = new RenameEventHandler(model, messageBox);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Control) )
            {
                handler = new RenameControlHandler(model, messageBox);
            }
            else
            {
                handler = new RenameDefaultHandler(model, messageBox);
            }
            return handler;
        }
    }
}
