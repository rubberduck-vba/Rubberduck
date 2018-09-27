using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.Interaction;
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
        private const string AppendUnderscoreFormat = "{0}_";
        private const string PrependUnderscoreFormat = "_{0}";

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
            var activeCodePane = _vbe.ActiveCodePane;
            _initialSelection = new Tuple<ICodePane, Selection>(activeCodePane, activeCodePane.IsWrappingNullReference ? Selection.Empty : activeCodePane.Selection);
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

            if (!(inputTarget is IInterfaceExposable))
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

            if (!ReferenceEquals(inputTarget, _model.Target))
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
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_TitleText) ?? false;

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
                var component = _state.ProjectsProvider.Component(target.QualifiedName.QualifiedModuleName);
                using (var controls = component.Controls)
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
                var component = _state.ProjectsProvider.Component(target.QualifiedName.QualifiedModuleName);
                using (var module = component.CodeModule)
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

            var conflicts = _state.DeclarationFinder.FindNewDeclarationNameConflicts(_model.NewName, _model.Target);

            if (conflicts.Any())
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName, _model.Target.IdentifierName);

                return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_Caption) ?? false;
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
                .Where(varDec => varDec.IsWithEvents).ToList();

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
            var interfaceMember = userTarget is IInterfaceExposable member && member.IsInterfaceMember
                ? userTarget
                : (userTarget as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

            IsInterfaceMemberRename = interfaceMember != null;
            return interfaceMember;
        }

        private void Rename()
        {
            Debug.Assert(!_model.NewName.Equals(_model.Target.IdentifierName, StringComparison.InvariantCultureIgnoreCase),
                            $"input validation fail: New Name equals Original Name ({_model.Target.IdentifierName})");

            var actionKeys = _renameActions.Keys.Where(decType => _model.Target.DeclarationType.HasFlag(decType)).ToList();
            if (actionKeys.Any())
            {
                Debug.Assert(actionKeys.Count == 1, $"{actionKeys.Count} Rename Actions have flag '{_model.Target.DeclarationType.ToString()}'");
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

            if (!IsInterfaceMemberRename)
            {
                return;
            }

            var implementations = _state.DeclarationFinder.FindAllInterfaceImplementingMembers()
                .Where(impl => ReferenceEquals(_model.Target.ParentDeclaration, impl.InterfaceImplemented)
                               && impl.InterfaceMemberImplemented.IdentifierName.Equals(_model.Target.IdentifierName));

            RenameDefinedFormatMembers(implementations.ToList(), PrependUnderscoreFormat);
        }

        private void RenameParameter()
        {
            if (_model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var parameters = _state.DeclarationFinder.MatchName(_model.Target.IdentifierName).Where(param =>
                   param.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                   && param.DeclarationType == DeclarationType.Parameter
                    && param.ParentDeclaration.IdentifierName.Equals(_model.Target.ParentDeclaration.IdentifierName)
                    && param.ParentDeclaration.ParentScopeDeclaration.Equals(_model.Target.ParentDeclaration.ParentScopeDeclaration));

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
            RenameDefinedFormatMembers(eventHandlers.ToList(), PrependUnderscoreFormat);
        }

        private void RenameVariable()
        {
            if ((_model.Target.Accessibility == Accessibility.Public ||
                 _model.Target.Accessibility == Accessibility.Implicit)
                && _model.Target.ParentDeclaration is ClassModuleDeclaration classDeclaration
                && classDeclaration.Subtypes.Any())
            {
                RenameMember();
            }
            else if (_model.Target.DeclarationType.HasFlag(DeclarationType.Control))
            {
                var component = _state.ProjectsProvider.Component(_model.Target.QualifiedName.QualifiedModuleName);
                using (var controls = component.Controls)
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
                RenameDefinedFormatMembers(controlEventHandlers.ToList(), AppendUnderscoreFormat);
            }
            else
            {
                RenameStandardElements(_model.Target, _model.NewName);
                if (_model.Target.IsWithEvents)
                {
                    var eventHandlers = _state.DeclarationFinder.FindHandlersForWithEventsField(_model.Target);
                    RenameDefinedFormatMembers(eventHandlers.ToList(), AppendUnderscoreFormat);
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
                        RenameDefinedFormatMembers(_state.DeclarationFinder.FindInterfaceMembersForImplementsContext(ctxt).ToList(), AppendUnderscoreFormat);
                    }
                }
            }

            var component = _state.ProjectsProvider.Component(_model.Target.QualifiedName.QualifiedModuleName);
            switch (component.Type)
            {
                case ComponentType.Document:
                    {
                        var properties = component.Properties;
                        var property = properties["_CodeName"];
                        {
                            property.Value = _model.NewName;
                        }
                        break;
                    }
                case ComponentType.UserForm:
                case ComponentType.VBForm:
                case ComponentType.MDIForm:
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
                        break;
                    }
                default:
                    {
                        if (_vbe.Kind == VBEKind.Hosted)
                        {
                            // VBA - rename code module
                            using (var codeModule = component.CodeModule)
                            {
                                Debug.Assert(!codeModule.IsWrappingNullReference, "input validation fail: Attempting to rename an ICodeModule wrapping a null reference");
                                codeModule.Name = _model.NewName;
                            }
                        }
                        else
                        {
                            // VB6 - rename component
                            component.Name = _model.NewName;
                        }
                        break;
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

        private void RenameDefinedFormatMembers(IReadOnlyCollection<Declaration> members, string underscoreFormat)
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
                .Where(reference =>
                    reference.Context.GetText() != "Me").GroupBy(r => r.QualifiedModuleName);

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
            var component = _state.ProjectsProvider.Component(qSelection.QualifiedName);
            using (var codeModule = component.CodeModule)
            {
                using (var codePane = codeModule.CodePane)
                {
                    if (!codePane.IsWrappingNullReference)
                    {
                        _initialSelection = new Tuple<ICodePane, Selection>(codePane, codePane.Selection);
                    }
                }
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
            _messageBox?.NotifyWarn(errorMsg, RubberduckUI.RenameDialog_Caption);
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

