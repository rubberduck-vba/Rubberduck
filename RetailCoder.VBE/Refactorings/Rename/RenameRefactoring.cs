using System.Linq;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;

namespace Rubberduck.Refactorings.Rename
{
    internal interface IRenameRefactoringHandler
    {
        void Rename();
        string ErrorMessage { get; }
    }

    public class RenameRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private RenameModel _model;

        public RenameRefactoring(IVBE vbe, IRefactoringPresenterFactory<IRenamePresenter> factory, IMessageBox messageBox, RubberduckParserState state)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
            _state = state;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            _model = presenter.Show();

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

        //The User has selected something in a code pane
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

        //The refactor starts from a Declaration only...not sure where the user selection is(?)
        public void Refactor(Declaration target)
        {
            if (!target.IsUserDefined)
            {
                //Todo: Add format string to UI resources
                var format = "Unable to rename Built-In declarations like '{0}'";
                var message = string.Format(format, target.IdentifierName);
                _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            var presenter = _factory.Create();
            _model = presenter.Show(target);

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

        private DialogResult ObtainApprovalToUseConflictingName(Declaration declaration)
        {
            var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                declaration);
            return _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation);
        }

        private bool UserCancelsRename()
        {
            var declarations = _state.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(_model.Target)
                .Where(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase));//.FirstOrDefault();

            if (declarations.Any())
            {
                var rename = ObtainApprovalToUseConflictingName(declarations.FirstOrDefault());

                if (rename == DialogResult.No)
                {
                    return true;
                }
            }
            return false;
        }

        private void Rename()
        {
            if (_model == null || _model.Declarations == null || _model.Target == null) { return; }

            if (UserCancelsRename()) { return; }

            var handler = GetHandler(_model);

            try
            {
                handler.Rename();
            }
            catch (Exception)
            {
                _messageBox.Show(handler.ErrorMessage, RubberduckUI.RenameDialog_Caption);
            }
        }

        //todo: Factory?
        private static IRenameRefactoringHandler GetHandler(RenameModel model)
        {
            IRenameRefactoringHandler handler = null;
            if (model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenamePropertyHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Parameter
                        && model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenamePropertyParameterHandler(model);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                handler = new RenameModuleHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Project)
            {
                handler = new RenameProjectHandler(model);
            }
            else if (model.Declarations.FindInterfaceMembers().Contains(model.Target))
            {
                handler = new RenameInterfaceHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Event || IsEvent(model))
            {
                handler = new RenameEventHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Control || IsControl(model))
            {
                handler = new RenameControlHandler(model);
            }
            else
            {
                handler = new RenameDefaultHandler(model);
            }
            return handler;
        }

        //When the user selects a control event to accomplish the rename, the declaration type is 
        //a procedure...so there is more work to do to determine if the procedure is a control event
        private static bool IsControl(RenameModel model)
        {
            var declaration = GetDeclarationOfType(model, DeclarationType.Control, false);
            return (declaration != null);
        }

        //Same reasoning for IsControl above
        private static bool IsEvent(RenameModel model)
        {
            var declaration = GetDeclarationOfType(model, DeclarationType.Event, true);
            return (declaration != null);
        }

        private static Declaration GetDeclarationOfType(RenameModel model, DeclarationType goalType, bool identifierIsOnRHS)
        {
            var declarationName = model.Target.IdentifierName;
            try
            {
                if (model.Target.IdentifierName.Contains("_"))
                {
                    if (identifierIsOnRHS)
                    {
                        //Events
                        declarationName = model.Target.IdentifierName.Remove(0, model.Target.IdentifierName.IndexOf("_") + 1);
                    }
                    else
                    {
                        //Controls
                        declarationName = model.Target.IdentifierName.Remove(model.Target.IdentifierName.IndexOf("_"));
                    }
                }
                var declaration = model.Declarations.Single(item =>
                    item.DeclarationType == goalType
                    && (declarationName.Equals(item.IdentifierName)));
                return declaration;
            }
            catch (Exception)
            {
                return null;
            }
        }

        //todo: decide if the following internal classes should become full-fledged independent classes

        internal class RenameDefaultHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameDefaultHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_DefaultRenameError"; } }

            public void Rename()
            {
                RenameUsages(_model, _model.Target);
                RewriteContent(_model);

                RenameDeclaration(_model, _model.Target);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenamePropertyHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenamePropertyHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_PropertyRenameError"; } }


            public void Rename()
            {
                // properties can have more than 1 member.
                var members = _model.Declarations.Named(_model.Target.IdentifierName)
                    .Where(item => item.ProjectId == _model.Target.ProjectId
                        && item.ComponentName == _model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property)).ToList();

                members.ForEach(member => RenameUsages(_model, member));
                RewriteContent(_model);

                RenameDeclaration(_model, _model.Target);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenamePropertyParameterHandler : IRenameRefactoringHandler
        {

            private RenameModel _model;

            public RenamePropertyParameterHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_PropertyParameterRenameError"; } }

            public void Rename()
            {
                var parameters = _model.Declarations.Where(d =>
                    d.DeclarationType == DeclarationType.Parameter
                    && d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                    && d.IdentifierName == _model.Target.IdentifierName).ToList();

                parameters.ForEach(param => RenameUsages(_model, param));
                RewriteContent(_model);

                parameters.ForEach(param => RenameDeclaration(_model, param));
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameInterfaceHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameInterfaceHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_InterfaceRenameError"; } }

            public void Rename()
            {
                RenameUsages(_model, _model.Target);
                RewriteContent(_model);

                var implementations = _model.Declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == _model.Target.ComponentName + '_' + _model.Target.IdentifierName)
                        .OrderByDescending(m => m.Selection.StartColumn).ToList();

                var newMemberName = _model.Target.ComponentName + '_' + _model.NewName;
                implementations.ForEach(imp => RenameDeclaration(_model, imp, newMemberName));
                RewriteContent(_model);

                RenameDeclaration(_model, _model.Target);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameEventHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameEventHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_EventRenameError"; } }

            public void Rename()
            {
                var eventDeclaration = GetDeclarationOfType(_model, DeclarationType.Event, true);
                if (null == eventDeclaration) { return; }

                var newEventName = _model.NewName;
                if (newEventName.Contains("_"))
                {
                    newEventName = _model.NewName.Remove(0, _model.NewName.IndexOf("_") + 1);
                }

                RenameUsages(_model, eventDeclaration, newEventName);
                RewriteContent(_model);

                var handlers = _model.Declarations.FindHandlersForEvent(eventDeclaration).ToList();
                handlers.ForEach(handler => RenameDeclaration(_model, handler.Item2, handler.Item1.IdentifierName + '_' + newEventName));
                RewriteContent(_model);

                RenameDeclaration(_model, eventDeclaration, newEventName);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameControlHandler : IRenameRefactoringHandler
        {
            private Declaration _target;
            private RenameModel _model;

            public RenameControlHandler(RenameModel model)
            {
                _target = model.Target;
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_ControlRenameError"; } }

            public void Rename()
            {
                var controlDeclaration = GetDeclarationOfType(_model, DeclarationType.Control, false);
                if (null == controlDeclaration) { return; }

                var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                var component = module.Parent;
                var control = component.Controls.SingleOrDefault(item => item.Name == controlDeclaration.IdentifierName);

                if (control == null) { return; }

                var newControlName = _model.NewName;
                if (newControlName.Contains("_"))
                {
                    newControlName = _model.NewName.Remove(_model.NewName.IndexOf("_"));
                }

                //User attempted to rename an event sub signature e.g. btn1_Click() to btn1_ClickAgain()
                if (newControlName.Equals(controlDeclaration.IdentifierName))
                {
                    //nothing to change
                    return;
                }

                RenameUsages(_model, controlDeclaration, newControlName);
                RewriteContent(_model);

                var handlers = _model.Declarations.FindEventHandlers(controlDeclaration).OrderByDescending(h => h.Selection.StartColumn).ToList();

                foreach (var handler in handlers)
                {
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', newControlName + '_');
                    RenameUsages(_model, handler, newMemberName);
                }
                RewriteContent(_model);

                foreach (var handler in handlers)
                {
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', newControlName + '_');
                    RenameDeclaration(_model, handler, newMemberName);
                }
                RewriteContent(_model);

                control.Name = newControlName;

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameProjectHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;
            public RenameProjectHandler(RenameModel model)
            {
                _model = model;
            }

            public string ErrorMessage { get { return RubberduckUI.RenameDialog_ProjectRenameError; } }

            public void Rename()
            {
                var projects = _model.VBE.VBProjects;
                var project = projects.SingleOrDefault(p => p.HelpFile == _model.Target.ProjectId);
                {
                    if (project != null)
                    {
                        project.Name = _model.NewName;
                    }
                }
            }
        }

        internal class RenameModuleHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameModuleHandler(RenameModel model)
            {
                _model = model;
            }

            public string ErrorMessage { get { return RubberduckUI.RenameDialog_ModuleRenameError; } }

            public void Rename()
            {
                RenameUsages(_model, _model.Target);
                RewriteContent(_model);

                var component = _model.Target.QualifiedName.QualifiedModuleName.Component;
                var module = component.CodeModule;
                {
                    if (module.IsWrappingNullReference)
                    {
                        return;
                    }

                    if (component.Type == ComponentType.Document)
                    {
                        var properties = component.Properties;
                        var property = properties["_CodeName"];
                        {
                            property.Value = _model.NewName;
                        }
                    }
                    else if (component.Type == ComponentType.UserForm)
                    {
                        var properties = component.Properties;
                        var property = properties["Caption"];
                        {
                            if ((string)property.Value == _model.Target.IdentifierName)
                            {
                                property.Value = _model.NewName;
                            }
                            component.Name = _model.NewName;
                        }
                    }
                    else
                    {
                        module.Name = _model.NewName;
                    }
                }
            }
        }

        private static void RenameUsages(RenameModel model, Declaration target)
        {
            RenameUsages(model, target, model.NewName);
        }

        private static void RenameUsages(RenameModel model, Declaration target, string newName)
        {
            var modules = target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                var rewriter = model.GetRewriter(grouping.Key);
                var module = grouping.Key.Component.CodeModule;
                {
                    foreach (var line in grouping.GroupBy(reference => reference.Selection.StartLine))
                    {
                        var lastSelection = Selection.Empty;
                        foreach (var reference in line.OrderByDescending(l => l.Selection.StartColumn))
                        {
                            if (reference.Selection == lastSelection)
                            {
                                continue;
                            }
                            var newContent = reference.Context.GetText().Replace(reference.IdentifierName, newName);
                            rewriter.Replace(reference.Context, newContent);
                            lastSelection = reference.Selection;
                        }
                    }
                }
            }
        }

        private static void RenameDeclaration(RenameModel model, Declaration target)
        {
            RenameDeclaration(model, target, model.NewName);
        }

        private static void RenameDeclaration(RenameModel model, Declaration target, string newName)
        {

            var component = target.QualifiedName.QualifiedModuleName.Component;
            var rewriter = model.GetRewriter(target);
            var module = component.CodeModule;
            {
                if (target.DeclarationType == DeclarationType.Parameter)
                {
                    var argList = (VBAParser.ArgListContext)target.Context.Parent;
                    var lineNum = argList.GetSelection().LineCount;

                    var newContent = target.Context.GetText().Replace(target.IdentifierName, newName);
                    rewriter.Replace(target, newContent);
                }
                else if (!target.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    var newContent = target.Context.GetText().Replace(target.IdentifierName, newName);
                    rewriter.Replace(target.Context, newContent);
                }
                else
                {
                    var members = model.Declarations.Named(target.IdentifierName)
                        .Where(item => item.ProjectId == target.ProjectId
                            && item.ComponentName == target.ComponentName
                            && item.DeclarationType.HasFlag(DeclarationType.Property));

                    foreach (var member in members)
                    {
                        var newContent = member.Context.GetText().Replace(member.IdentifierName, newName);
                        rewriter.Replace(member.Context, newContent);
                    }
                }
            }
        }

        private static void RewriteContent(RenameModel model)
        {
            model.Rewrite();
            model.ClearRewriters();
        }
    }
}
