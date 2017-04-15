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

            bool hasNullReferences;
            SetModelMember(presenter, out hasNullReferences);

            if(hasNullReferences) { return; }

            _model.Target = PreprocessSelectedTarget(_model.Target);

            ValidateConditionsForRename();

            if (_renameOperationIsCancelled) { return; }

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

            bool hasNullReferences;
            SetModelMember(presenter, out hasNullReferences);

            if (hasNullReferences) { return; }

            if (target == null)
            {
                _messageBox?.Show(RubberduckUI.RefactorRename_TargetNotDefinedError, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            if (!target.IsUserDefined)
            {
                if (_messageBox == null) { return; }
                var message = string.Format(RubberduckUI.RefactorRename_TargetNotUserDefinedError, target.QualifiedName);
                _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            _model.Target = PreprocessSelectedTarget(target);

            ValidateConditionsForRename();

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
            if(presenter == null)
            {
                hasNullReferences = true;
                return;
            }

            _model = presenter.Model;

            if (_model == null)
            {
                hasNullReferences = true;
                return;
            }

            if (null == _model.Target)
            {
                hasNullReferences = true;
                _messageBox?.Show(RubberduckUI.RefactorRename_TargetNotDefinedError, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        //For Controls, Events, and Interfaces - make sure what the user is presented with
        //the declaration rather than a handler or implementation so that there is no confusion about what is 
        //being changed.
        //(e.g., If the user selects a control event handler like 'bnt1_Click', 
        //he is presented with 'bnt1' to rename.  
        private Declaration PreprocessSelectedTarget(Declaration selectedTarget)
        {
            if(selectedTarget == null) { return null; }

            if (!selectedTarget.DeclarationType.HasFlag(DeclarationType.Procedure))
            {
                return selectedTarget;
            }

            Declaration control;
            if(IsControlEventHandler(selectedTarget, out control))
            {
                return control;
            }
            Debug.Assert(selectedTarget.DeclarationType.HasFlag(DeclarationType.Procedure), "Control Resolver Error: modified target type to " + selectedTarget.DeclarationType.ToString());

            Declaration eventDeclaration;
            if (IsUserEventRelated(selectedTarget, out eventDeclaration))
            {
                return eventDeclaration;
            }
            Debug.Assert(selectedTarget.DeclarationType.HasFlag(DeclarationType.Procedure), "Event Resolver Error: modified target type to " + selectedTarget.DeclarationType.ToString());

            Declaration interfaceDefinition;
            if(IsInterfaceImplementation(selectedTarget, out interfaceDefinition))
            {
                if (selectedTarget != interfaceDefinition)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, selectedTarget.IdentifierName, selectedTarget.ComponentName, selectedTarget.IdentifierName);

                    var confirm = _messageBox.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    _renameOperationIsCancelled = (confirm == DialogResult.No);
                }
                return interfaceDefinition;
            }
            Debug.Assert(selectedTarget.DeclarationType.HasFlag(DeclarationType.Procedure), "Interface Resolver Error: modified target type to " + selectedTarget.DeclarationType.ToString());

            return selectedTarget;
        }

        private void ValidateConditionsForRename()
        {
            //more checks once the RenameModel's target is resolved
            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Control))
            {
                var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                var component = module.Parent;
                var control = component.Controls.SingleOrDefault(item => item.Name == _model.Target.IdentifierName);

                if (control == null)
                {
                    _renameOperationIsCancelled = true;
                    var message = string.Format(RubberduckUI.RenameDialog_ControlRenameError, _model.Target.IdentifierName);
                    _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
                }
            }

            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                var component = _model.Target.QualifiedName.QualifiedModuleName.Component;
                var module = component.CodeModule;
                if (module.IsWrappingNullReference)
                {
                    _renameOperationIsCancelled = true;
                    var message = RubberduckUI.RenameDialog_ModuleRenameError;
                    _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
                }
            }

        }

        private bool UserCancelsDueToNameConflict()
        {
            var declarations = _state.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(_model.Target)
                .Where(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase));//.FirstOrDefault();

            if (declarations.Any())
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                    declarations.FirstOrDefault().IdentifierName);

                var rename = _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation);

                _renameOperationIsCancelled = (rename == DialogResult.No);
            }
            return _renameOperationIsCancelled;
        }

        private bool IsControlEventHandler(Declaration userTarget, out Declaration control)
        {
            control = null;
            var declarationsOfInterest = _model.State.AllUserDeclarations.Where(d => d.DeclarationType.HasFlag(DeclarationType.Control)
                    && userTarget.IdentifierName.StartsWith(d.IdentifierName));

            if(declarationsOfInterest.Any())
            {
                control = declarationsOfInterest.FirstOrDefault();
                return true;
            }
            return false;
        }

        private bool IsUserEventRelated(Declaration userTarget, out Declaration eventDeclaration)
        {
            var declarationsOfInterest = _model.State.AllUserDeclarations.Where(d => d.DeclarationType.HasFlag(DeclarationType.Event)
                    && userTarget.IdentifierName.Contains(d.IdentifierName));

            eventDeclaration = null;
            if (declarationsOfInterest.Any())
            {
                eventDeclaration = declarationsOfInterest.FirstOrDefault();
                return true;
            }
            return false;
        }

        private bool IsInterfaceImplementation(Declaration userTarget, out Declaration interfaceDefinition)
        {
            interfaceDefinition = null;

            var interfaceImplementation = _model.State.DeclarationFinder.FindAllInterfaceImplementingMembers()
                .SingleOrDefault(m => m.Equals(userTarget));

            if(interfaceImplementation == null)
            {
                interfaceDefinition = userTarget;
                return true;
            }

            var matches = _model.State.DeclarationFinder.FindAllInterfaceMembers()
                        .Where(m => m.IsUserDefined && interfaceImplementation.IdentifierName == m.ComponentName + '_' + m.IdentifierName).ToList();

            var interfaceMember =  matches.Count > 1
                ? matches.SingleOrDefault(m => m.ProjectId == interfaceImplementation.ProjectId)
                : matches.FirstOrDefault();

            if (interfaceMember != null)
            {
                interfaceDefinition = interfaceMember;
                return true;
            }
            return false;
        }

        private void Rename()
        {
            var handler = GetHandler(_model);
            try
            {
                handler.Rename(_model.Target, _model.NewName);
                if (handler.RequestParseAfterRename)
                {
                    _model.State.OnParseRequested(this);
                }
            }
            catch (RuntimeBinderException)
            {
                _messageBox?.Show(handler.ErrorMessage, RubberduckUI.RenameDialog_Caption);
            }
            catch (COMException)
            {
                _messageBox?.Show(handler.ErrorMessage, RubberduckUI.RenameDialog_Caption);
            }
        }

        private IRename GetHandler(RenameModel model)
        {
            IRename handler;
            if (model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenameRefactorProperty(model.State);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Parameter)
                        && model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenameRefactorPropertyParameter(model.State);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                handler = new RenameRefactorModule(model.State);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Project))
            {
                handler = new RenameRefactorProject(model);
            }
            else if (model.Declarations.FindInterfaceMembers().Contains(model.Target))
            {
                handler = new RenameRefactorInterface(model.State);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Event)) 
            {
                handler = new RenameRefactorEvent(model.State);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Control) )
            {
                handler = new RenameRefactorControl(model.State);
            }
            else
            {
                handler = new RenameRefactorDefault(model.State);
            }
            return handler;
        }
    }
}
