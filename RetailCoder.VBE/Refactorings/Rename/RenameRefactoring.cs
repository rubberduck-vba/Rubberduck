using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Diagnostics;
using Microsoft.CSharp.RuntimeBinder;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : IRefactoring
    {
        private const string _appendUnderscoreFormat = "{0}_";
        private const string _prependUnderscoreFormat = "_{0}";

        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private RenameModel _model;
        private Tuple<ICodePane, Selection> _initialSelection;
        private readonly List<QualifiedModuleName> _modulesToRewrite;
        private readonly Dictionary<DeclarationType, Action> _renameActions;
        private readonly List<string> _neverRenameIdentifiers;

        private bool IsInterfaceMemberRename { set; get; }
        private bool IsControlEventHandlerRename { set; get; }
        private bool IsUserEventHandlerRename { set; get; }
        private bool RequestParseAfterRename { set; get; }

        public RenameRefactoring(IVBE vbe, IRefactoringPresenterFactory<IRenamePresenter> factory, IMessageBox messageBox, RubberduckParserState state)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
            _state = state;
            _model = null;
            _initialSelection = new Tuple<ICodePane, Selection>(_vbe.ActiveCodePane, _vbe.ActiveCodePane.IsWrappingNullReference ? Selection.Empty : _vbe.ActiveCodePane.Selection);
            _modulesToRewrite = new List<QualifiedModuleName>();
            _renameActions = new Dictionary<DeclarationType, Action>
            {
                {DeclarationType.Member, RenameMember},
                {DeclarationType.Parameter, RenameParameter},
                {DeclarationType.Event, RenameEvent},
                {DeclarationType.Variable, RenameVariable},
                {DeclarationType.Module, RenameModule},
                {DeclarationType.Project, RenameProject}
            };
            IsInterfaceMemberRename = false;
            RequestParseAfterRename = true;
            _neverRenameIdentifiers = NeverRenameList();
        }

        public void Refactor(QualifiedSelection qualifiedSelection)
        {
            CacheInitialSelection(qualifiedSelection);
            Refactor();
        }

        public void Refactor()
        {
            var presenter = CreateRenamePresenter();
            if (presenter != null)
            {
                RefactorImpl(presenter.Model.Target, presenter);
                RestoreInitialSelection();
            }
        }

        public void Refactor(Declaration target)
        {
            var presenter = CreateRenamePresenter();
            if (presenter != null)
            {
                RefactorImpl(target, presenter);
                RestoreInitialSelection();
            }
        }

        private void RefactorImpl(Declaration inputTarget, IRenamePresenter presenter)
        {
            try
            {
                if (!TrySetRenameTargetFromInputTarget(inputTarget))
                {
                    return;
                }

                if (TrySetNewName(presenter))
                {
                    Rename();
                    Rewrite();
                    Reparse();
                }
            }
            catch (RuntimeBinderException rbEx)
            {
                PresentRenameErrorMessage($"{BuildDefaultErrorMessage(_model.Target)}: {rbEx.Message}");
            }
            catch (COMException comEx)
            {
                PresentRenameErrorMessage($"{BuildDefaultErrorMessage(_model.Target)}: {comEx.Message}");
            }
            catch (Exception unhandledEx)
            {
                PresentRenameErrorMessage($"{BuildDefaultErrorMessage(_model.Target)}: {unhandledEx.Message}");
            }
        }

        private IRenamePresenter CreateRenamePresenter()
        {
            var presenter = _factory.Create();
            if (presenter != null)
            {
                _model = presenter.Model;
            }
            if (presenter == null || _model == null)
            {
                PresentRenameErrorMessage(RubberduckUI.RefactorRename_TargetNotDefinedError);
                return null;
            }
            return presenter;
        }

        private bool TrySetRenameTargetFromInputTarget(Declaration inputTarget)
        {
            if (!IsValidTarget(inputTarget)) { return false; }

            if (!inputTarget.DeclarationType.HasFlag(DeclarationType.Member))
            {
                _model.Target = inputTarget;
                return true;
            }

            if (_neverRenameIdentifiers.Contains(inputTarget.IdentifierName))
            {
                PresentRenameErrorMessage(string.Format(RubberduckUI.RenameDialog_BuiltInNameError, $"{inputTarget.ComponentName}: {inputTarget.DeclarationType}", inputTarget.IdentifierName));
                return false;
            }

            _model.Target = ResolveRenameTargetIfEventHandlerSelected(inputTarget) ??
                            ResolveRenameTargetIfInterfaceImplementationSelected(inputTarget) ??
                            inputTarget;

            if (inputTarget != _model.Target)
            {
                //Resolved to a target other than the input target selected by the user.
                //Check that the resolved target is valid and that the user wants to continue with the rename 
                if (!IsValidTarget(_model.Target)) { return false; }

                if (IsControlEventHandlerRename)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsControlEventHandler, inputTarget.IdentifierName, _model.Target.DeclarationType, _model.Target.IdentifierName);
                    return UserConfirmsRenameOfResolvedTarget(message);
                }

                if (IsUserEventHandlerRename)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsEventHandlerImplementation, inputTarget.IdentifierName, _model.Target.DeclarationType, _model.Target.IdentifierName);
                    return UserConfirmsRenameOfResolvedTarget(message);
                }

                if (IsInterfaceMemberRename)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, inputTarget.IdentifierName, _model.Target.ComponentName, _model.Target.IdentifierName);
                    return UserConfirmsRenameOfResolvedTarget(message);
                }

                Debug.Assert(false, $"Resolved rename target ({_model.Target.Scope}) unhandled");
            }
            return true;
        }

        private bool UserConfirmsRenameOfResolvedTarget(string message)
        {
            var confirm = _messageBox?.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            return confirm.HasValue && confirm.Value == DialogResult.Yes;
        }

        private Declaration ResolveRenameTargetIfEventHandlerSelected(Declaration selectedTarget)
        {
            if (selectedTarget.DeclarationType.HasFlag(DeclarationType.Procedure) && selectedTarget.IdentifierName.Contains("_"))
            {
                return ResolveEventHandlerToControl(selectedTarget) ??
                       ResolveEventHandlerToUserEvent(selectedTarget);
            }
            return null;
        }

        private bool IsValidTarget(Declaration target)
        {
            if (target == null)
            {
                PresentRenameErrorMessage(RubberduckUI.RefactorRename_TargetNotDefinedError);
                return false;
            }

            if (!target.IsUserDefined)
            {
                PresentRenameErrorMessage(string.Format(RubberduckUI.RefactorRename_TargetNotUserDefinedError, target.QualifiedName));
                return false;
            }

            if (target.DeclarationType.HasFlag(DeclarationType.Control))
            {
                using (var controls = target.QualifiedName.QualifiedModuleName.Component.Controls)
                {
                    using (var control = controls.FirstOrDefault(item => item.Name == target.IdentifierName))
                    {
                        if (control == null)
                        {
                            PresentRenameErrorMessage($"{BuildDefaultErrorMessage(target)} - Null control reference");
                            return false;
                        }
                    }
                }
            }
            else if (target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                using (var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule)
                {
                    if (module.IsWrappingNullReference)
                    {
                        PresentRenameErrorMessage($"{BuildDefaultErrorMessage(target)} - Null Module reference");
                        return false;
                    }
                }
            }
            return true;
        }

        private bool TrySetNewName(IRenamePresenter presenter)
        {
            var result = presenter.Show(_model.Target);
            if (result == null) { return false; }

            var conflictDeclarations = _state.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(_model.Target)
                .Where(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase));

            if (conflictDeclarations.Any())
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                    conflictDeclarations.FirstOrDefault().IdentifierName);

                var rename = _messageBox?.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation);

                return rename.HasValue && rename.Value == DialogResult.Yes;
            }

            return true;
        }

        private Declaration ResolveEventHandlerToControl(Declaration userTarget)
        {
            var control = _state.DeclarationFinder.UserDeclarations(DeclarationType.Control)
                .FirstOrDefault(ctrl => userTarget.Scope.StartsWith($"{ctrl.ParentScope}.{ctrl.IdentifierName}_"));

            IsControlEventHandlerRename = control != null;

            return FindEventHandlersForControl(control).Contains(userTarget) ? control : null;
        }

        private Declaration ResolveEventHandlerToUserEvent(Declaration userTarget)
        {
            var withEventsDeclarations = _state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(varDec => varDec.IsWithEvents);

            if (!withEventsDeclarations.Any()) { return null; }

            foreach (var withEvent in withEventsDeclarations)
            {
                if (userTarget.IdentifierName.StartsWith($"{withEvent.IdentifierName}_"))
                {
                    if (_state.DeclarationFinder.FindHandlersForWithEventsField(withEvent).Contains(userTarget))
                    {
                        var eventName = userTarget.IdentifierName.Remove(0, $"{withEvent.IdentifierName}_".Length);

                        var eventDeclaration = _state.DeclarationFinder.UserDeclarations(DeclarationType.Event).FirstOrDefault(ev => ev.IdentifierName.Equals(eventName)
                                && withEvent.AsTypeName.Equals(ev.ParentDeclaration.IdentifierName));

                        IsUserEventHandlerRename = eventDeclaration != null;

                        return eventDeclaration;
                    }
                }
            }
            return null;
        }

        private Declaration ResolveRenameTargetIfInterfaceImplementationSelected(Declaration userTarget)
        {
            var interfaceMember = _state.DeclarationFinder.FindAllInterfaceMembers()
                .Where(member => member.Equals(userTarget)
                    || (member.ProjectId.Equals(userTarget.ProjectId)
                        && member.DeclarationType == userTarget.DeclarationType
                        && $"{member.ParentDeclaration.IdentifierName}_{member.IdentifierName}".Equals(userTarget.IdentifierName))).FirstOrDefault();

            IsInterfaceMemberRename = interfaceMember != null;
            return interfaceMember;
        }

        private void Rename()
        {
            Debug.Assert(!_model.NewName.Equals(_model.Target.IdentifierName, StringComparison.InvariantCultureIgnoreCase),
                            $"input validation fail: New Name equals Original Name ({_model.Target.IdentifierName})");

            var actionKeys = _renameActions.Keys.Where(decType => _model.Target.DeclarationType.HasFlag(decType));
            if (actionKeys.Any())
            {
                Debug.Assert(actionKeys.Count() == 1, $"{actionKeys.Count()} Rename Actions have flag '{_model.Target.DeclarationType.ToString()}'");
                _renameActions[actionKeys.FirstOrDefault()]();
            }
            else
            {
                RenameStandardElements(_model.Target, _model.NewName);
            }
        }

        private void RenameMember()
        {
            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var members = _state.DeclarationFinder.MatchName(_model.Target.IdentifierName)
                    .Where(item => item.ProjectId == _model.Target.ProjectId
                        && item.ComponentName == _model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    RenameStandardElements(member, _model.NewName);
                }
            }
            else
            {
                RenameStandardElements(_model.Target, _model.NewName);
            }

            if (IsInterfaceMemberRename)
            {
                var implementations = _state.DeclarationFinder.FindAllInterfaceImplementingMembers()
                    .Where(member => member.ProjectId == _model.Target.ProjectId
                        && member.IdentifierName.Equals($"{_model.Target.ComponentName}_{_model.Target.IdentifierName}"));

                RenameDefinedFormatMembers(implementations, _prependUnderscoreFormat);
            }
        }

        private void RenameParameter()
        {
            if (_model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var parameters = _state.DeclarationFinder.MatchName(_model.Target.IdentifierName).Where(param =>
                   param.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                   && param.DeclarationType == DeclarationType.Parameter);

                foreach (var param in parameters)
                {
                    RenameStandardElements(param, _model.NewName);
                }
            }
            else
            {
                RenameStandardElements(_model.Target, _model.NewName);
            }
        }

        private void RenameEvent()
        {
            RenameStandardElements(_model.Target, _model.NewName);

            var withEventsDeclarations = _state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(varDec => varDec.IsWithEvents && varDec.AsTypeName.Equals(_model.Target.ParentDeclaration.IdentifierName));

            var eventHandlers = withEventsDeclarations.SelectMany(we => _state.DeclarationFinder.FindHandlersForWithEventsField(we));
            RenameDefinedFormatMembers(eventHandlers, _prependUnderscoreFormat);
        }

        private void RenameVariable()
        {
            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Control))
            {
                using (var controls = _model.Target.QualifiedName.QualifiedModuleName.Component.Controls)
                {
                    using (var control = controls.SingleOrDefault(item => item.Name == _model.Target.IdentifierName))
                    {
                        Debug.Assert(control != null,
                            $"input validation fail: unable to locate '{_model.Target.IdentifierName}' in Controls collection");

                        control.Name = _model.NewName;
                    }
                }
                RenameReferences(_model.Target, _model.NewName);
                var controlEventHandlers = FindEventHandlersForControl(_model.Target);
                RenameDefinedFormatMembers(controlEventHandlers, _appendUnderscoreFormat);
            }
            else
            {
                RenameStandardElements(_model.Target, _model.NewName);
                if (_model.Target.IsWithEvents)
                {
                    var eventHandlers = _state.DeclarationFinder.FindHandlersForWithEventsField(_model.Target);
                    RenameDefinedFormatMembers(eventHandlers, _appendUnderscoreFormat);
                }
            }
        }

        private void RenameModule()
        {
            RequestParseAfterRename = false;

            RenameReferences(_model.Target, _model.NewName);

            if (_model.Target.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                foreach (var reference in _model.Target.References)
                {
                    var ctxt = reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>();
                    if (ctxt != null)
                    {
                        RenameDefinedFormatMembers(_state.DeclarationFinder.FindInterfaceMembersForImplementsContext(ctxt), _appendUnderscoreFormat);
                    }
                }
            }

            var component = _model.Target.QualifiedName.QualifiedModuleName.Component;
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
                using (var codeModule = component.CodeModule)
                {
                    Debug.Assert(!codeModule.IsWrappingNullReference, "input validation fail: Attempting to rename an ICodeModule wrapping a null reference");
                    codeModule.Name = _model.NewName;
                }
            }
        }

        private void RenameProject()
        {
            RequestParseAfterRename = false;
            var project = ProjectById(_vbe, _model.Target.ProjectId);

            if (project != null)
            {
                project.Name = _model.NewName;
                project.Dispose();
            }
        }

        private IVBProject ProjectById(IVBE vbe, string projectId)
        {
            using (var projects = vbe.VBProjects)
            {
                foreach (var project in projects)
                {
                    if (project.ProjectId == projectId)
                    {
                        return project;
                    }
                    project.Dispose();
                }
            }
            return null;
        }

        private void RenameDefinedFormatMembers(IEnumerable<Declaration> members, string underscoreFormat)
        {
            if (!members.Any()) { return; }

            var targetFragment = string.Format(underscoreFormat, _model.Target.IdentifierName);
            var replacementFragment = string.Format(underscoreFormat, _model.NewName);
            foreach (var member in members)
            {
                var newMemberName = member.IdentifierName.Replace(targetFragment, replacementFragment);
                RenameStandardElements(member, newMemberName);
            }
        }

        private void RenameStandardElements(Declaration target, string newName)
        {
            RenameReferences(target, newName);
            RenameDeclaration(target, newName);
        }

        private void RenameReferences(Declaration target, string newName)
        {
            var modules = target.References
                .Where(reference => reference.Context.GetText() != "Me")
                .GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                _modulesToRewrite.Add(grouping.Key);
                var rewriter = _state.GetRewriter(grouping.Key);
                foreach (var reference in grouping)
                {
                    rewriter.Replace(reference.Context, newName);
                }
            }
        }

        private void RenameDeclaration(Declaration target, string newName)
        {
            _modulesToRewrite.Add(target.QualifiedName.QualifiedModuleName);
            var rewriter = _state.GetRewriter(target.QualifiedName.QualifiedModuleName);

            if (target.Context is IIdentifierContext context)
            {
                rewriter.Replace(context.IdentifierTokens, newName);
            }
        }

        private void Rewrite()
        {
            foreach (var module in _modulesToRewrite.Distinct())
            {
                _state.GetRewriter(module).Rewrite();
            }
        }

        private void Reparse()
        {
            if (RequestParseAfterRename)
            {
                _state.OnParseRequested(this);
            }
        }

        private IEnumerable<Declaration> FindEventHandlersForControl(Declaration control)
        {
            if (control != null && control.DeclarationType.HasFlag(DeclarationType.Control))
            {
                return _state.DeclarationFinder.FindEventHandlers()
                    .Where(ev => ev.Scope.StartsWith($"{control.ParentScope}.{control.IdentifierName}_"));
            }
            return Enumerable.Empty<Declaration>();
        }

        private void CacheInitialSelection(QualifiedSelection qSelection)
        {
            if (!qSelection.QualifiedName.Component.CodeModule.CodePane.IsWrappingNullReference)
            {
                _initialSelection = new Tuple<ICodePane, Selection>(qSelection.QualifiedName.Component.CodeModule.CodePane, qSelection.QualifiedName.Component.CodeModule.CodePane.Selection);
            }
        }

        private void RestoreInitialSelection()
        {
            if (!_initialSelection.Item1.IsWrappingNullReference)
            {
                _initialSelection.Item1.Selection = _initialSelection.Item2;
            }
        }

        private void PresentRenameErrorMessage(string errorMsg)
        {
            _messageBox?.Show(errorMsg, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private string BuildDefaultErrorMessage(Declaration target)
        {
            var messageFormat = IsInterfaceMemberRename ? RubberduckUI.RenameDialog_InterfaceRenameError : RubberduckUI.RenameDialog_DefaultRenameError;
            return string.Format(messageFormat, target.DeclarationType.ToString(), target.IdentifierName);
        }

        private List<string> NeverRenameList()
        {
            return _state.DeclarationFinder.FindEventHandlers()
                    .Where(ev => ev.IdentifierName.StartsWith("Class_")
                            || ev.IdentifierName.StartsWith("UserForm_")
                            || ev.IdentifierName.StartsWith("auto_"))
                    .Select(dec => dec.IdentifierName).ToList();
        }
    }
}

