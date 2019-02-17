using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.Interaction;
using Rubberduck.VBEditor;
using System;
using System.Diagnostics;
using Microsoft.CSharp.RuntimeBinder;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : InteractiveRefactoringBase<IRenamePresenter, RenameModel>
    {
        private const string AppendUnderscoreFormat = "{0}_";
        private const string PrependUnderscoreFormat = "_{0}";

        private readonly IMessageBox _messageBox;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IDictionary<DeclarationType, Action<IRewriteSession>> _renameActions;
        private readonly List<string> _neverRenameIdentifiers;

        private bool IsInterfaceMemberRename { set; get; }
        private bool IsControlEventHandlerRename { set; get; }
        private bool IsUserEventHandlerRename { set; get; }
        public RenameRefactoring(IRefactoringPresenterFactory factory, IMessageBox messageBox, IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _messageBox = messageBox;
            _declarationFinderProvider = declarationFinderProvider;
            _projectsProvider = projectsProvider;
            Model = null;

            _renameActions = new Dictionary<DeclarationType, Action<IRewriteSession>>
            {
                {DeclarationType.Member, RenameMember},
                {DeclarationType.Parameter, RenameParameter},
                {DeclarationType.Event, RenameEvent},
                {DeclarationType.Variable, RenameVariable},
                {DeclarationType.Module, RenameModule},
                {DeclarationType.Project, RenameProject}
            };
            IsInterfaceMemberRename = false;
            _neverRenameIdentifiers = NeverRenameList();
        }

        public override void Refactor(QualifiedSelection target)
        {
            Refactor(InitializeModel(target));
        }

        private RenameModel InitializeModel(QualifiedSelection targetSelection)
        {
            return new RenameModel(_declarationFinderProvider.DeclarationFinder, targetSelection);
        }

        protected override void RefactorImpl(IRenamePresenter presenter)
        {
            RefactorImpl(Model.Target, presenter);
        }

        public override void Refactor(Declaration target)
        {
            Refactor(InitializeModel(target));
        }

        private RenameModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                return null;
            }

            return InitializeModel(target.QualifiedSelection);
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
                    var rewriteSession = RewritingManager.CheckOutCodePaneSession();
                    Rename(rewriteSession);
                    rewriteSession.TryRewrite();
                }
            }
            catch (RuntimeBinderException rbEx)
            {
                PresentRenameErrorMessage($"{BuildDefaultErrorMessage(Model.Target)}: {rbEx.Message}");
            }
            catch (COMException comEx)
            {
                PresentRenameErrorMessage($"{BuildDefaultErrorMessage(Model.Target)}: {comEx.Message}");
            }
            catch (Exception unhandledEx)
            {
                PresentRenameErrorMessage($"{BuildDefaultErrorMessage(Model.Target)}: {unhandledEx.Message}");
            }
        }

        private bool TrySetRenameTargetFromInputTarget(Declaration inputTarget)
        {
            if (!IsValidTarget(inputTarget)) { return false; }

            if (!(inputTarget is IInterfaceExposable))
            {
                Model.Target = inputTarget;
                return true;
            }

            if (_neverRenameIdentifiers.Contains(inputTarget.IdentifierName))
            {
                PresentRenameErrorMessage(string.Format(RubberduckUI.RenameDialog_BuiltInNameError, $"{inputTarget.ComponentName}: {inputTarget.DeclarationType}", inputTarget.IdentifierName));
                return false;
            }

            Model.Target = ResolveRenameTargetIfEventHandlerSelected(inputTarget) ??
                            ResolveRenameTargetIfInterfaceImplementationSelected(inputTarget) ??
                            inputTarget;

            if (!ReferenceEquals(inputTarget, Model.Target))
            {
                //Resolved to a target other than the input target selected by the user.
                //Check that the resolved target is valid and that the user wants to continue with the rename 
                if (!IsValidTarget(Model.Target)) { return false; }

                if (IsControlEventHandlerRename)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsControlEventHandler, inputTarget.IdentifierName, Model.Target.DeclarationType, Model.Target.IdentifierName);
                    return UserConfirmsRenameOfResolvedTarget(message);
                }

                if (IsUserEventHandlerRename)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsEventHandlerImplementation, inputTarget.IdentifierName, Model.Target.DeclarationType, Model.Target.IdentifierName);
                    return UserConfirmsRenameOfResolvedTarget(message);
                }

                if (IsInterfaceMemberRename)
                {
                    var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, inputTarget.IdentifierName, Model.Target.ComponentName, Model.Target.IdentifierName);
                    return UserConfirmsRenameOfResolvedTarget(message);
                }

                Debug.Assert(false, $"Resolved rename target ({Model.Target.Scope}) unhandled");
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
                var component = _projectsProvider.Component(target.QualifiedName.QualifiedModuleName);
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
                var component = _projectsProvider.Component(target.QualifiedName.QualifiedModuleName);
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
            var result = presenter.Show(Model.Target);
            if (result == null)
            {
                return false;
            }

            Model = result;

            var conflicts = _declarationFinderProvider.DeclarationFinder.FindNewDeclarationNameConflicts(Model.NewName, Model.Target);

            if (conflicts.Any())
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, Model.NewName, Model.Target.IdentifierName);

                return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_Caption) ?? false;
            }

            return true;
        }

        private Declaration ResolveEventHandlerToControl(Declaration userTarget)
        {
            var control = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Control)
                .FirstOrDefault(ctrl => userTarget.Scope.StartsWith($"{ctrl.ParentScope}.{ctrl.IdentifierName}_"));

            IsControlEventHandlerRename = control != null;

            return FindEventHandlersForControl(control).Contains(userTarget) ? control : null;
        }

        private Declaration ResolveEventHandlerToUserEvent(Declaration userTarget)
        {
            var withEventsDeclarations = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(varDec => varDec.IsWithEvents).ToList();

            if (!withEventsDeclarations.Any()) { return null; }

            foreach (var withEvent in withEventsDeclarations)
            {
                if (userTarget.IdentifierName.StartsWith($"{withEvent.IdentifierName}_"))
                {
                    if (_declarationFinderProvider.DeclarationFinder.FindHandlersForWithEventsField(withEvent).Contains(userTarget))
                    {
                        var eventName = userTarget.IdentifierName.Remove(0, $"{withEvent.IdentifierName}_".Length);

                        var eventDeclaration = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Event).FirstOrDefault(ev => ev.IdentifierName.Equals(eventName)
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

        private void Rename(IRewriteSession rewriteSession)
        {
            Debug.Assert(!Model.NewName.Equals(Model.Target.IdentifierName, StringComparison.InvariantCultureIgnoreCase),
                            $"input validation fail: New Name equals Original Name ({Model.Target.IdentifierName})");

            var actionKeys = _renameActions.Keys.Where(decType => Model.Target.DeclarationType.HasFlag(decType)).ToList();
            if (actionKeys.Any())
            {
                Debug.Assert(actionKeys.Count == 1, $"{actionKeys.Count} Rename Actions have flag '{Model.Target.DeclarationType.ToString()}'");
                _renameActions[actionKeys.FirstOrDefault()](rewriteSession);
            }
            else
            {
                RenameStandardElements(Model.Target, Model.NewName, rewriteSession);
            }
        }

        private void RenameMember(IRewriteSession rewriteSession)
        {
            if (Model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var members = _declarationFinderProvider.DeclarationFinder.MatchName(Model.Target.IdentifierName)
                    .Where(item => item.ProjectId == Model.Target.ProjectId
                        && item.ComponentName == Model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    RenameStandardElements(member, Model.NewName, rewriteSession);
                }
            }
            else
            {
                RenameStandardElements(Model.Target, Model.NewName, rewriteSession);
            }

            if (!IsInterfaceMemberRename)
            {
                return;
            }
            
            var implementations = _declarationFinderProvider.DeclarationFinder.FindAllInterfaceImplementingMembers()
                .Where(impl => ReferenceEquals(Model.Target.ParentDeclaration, impl.InterfaceImplemented)
                               && impl.InterfaceMemberImplemented.IdentifierName.Equals(Model.Target.IdentifierName));

            RenameDefinedFormatMembers(implementations.ToList(), PrependUnderscoreFormat, rewriteSession);
        }

        private void RenameParameter(IRewriteSession rewriteSession)
        {
            if (Model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var parameters = _declarationFinderProvider.DeclarationFinder.MatchName(Model.Target.IdentifierName).Where(param =>
                   param.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                   && param.DeclarationType == DeclarationType.Parameter
                    && param.ParentDeclaration.IdentifierName.Equals(Model.Target.ParentDeclaration.IdentifierName)
                    && param.ParentDeclaration.ParentScopeDeclaration.Equals(Model.Target.ParentDeclaration.ParentScopeDeclaration));

                foreach (var param in parameters)
                {
                    RenameStandardElements(param, Model.NewName, rewriteSession);
                }
            }
            else
            {
                RenameStandardElements(Model.Target, Model.NewName, rewriteSession);
            }
        }

        private void RenameEvent(IRewriteSession rewriteSession)
        {
            RenameStandardElements(Model.Target, Model.NewName, rewriteSession);

            var withEventsDeclarations = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(varDec => varDec.IsWithEvents && varDec.AsTypeName.Equals(Model.Target.ParentDeclaration.IdentifierName));

            var eventHandlers = withEventsDeclarations.SelectMany(we => _declarationFinderProvider.DeclarationFinder.FindHandlersForWithEventsField(we));
            RenameDefinedFormatMembers(eventHandlers.ToList(), PrependUnderscoreFormat, rewriteSession);
        }

        private void RenameVariable(IRewriteSession rewriteSession)
        {
            if ((Model.Target.Accessibility == Accessibility.Public ||
                 Model.Target.Accessibility == Accessibility.Implicit)
                && Model.Target.ParentDeclaration is ClassModuleDeclaration classDeclaration
                && classDeclaration.Subtypes.Any())
            {
                RenameMember(rewriteSession);
            }
            else if (Model.Target.DeclarationType.HasFlag(DeclarationType.Control))
            {
                var component = _projectsProvider.Component(Model.Target.QualifiedName.QualifiedModuleName);
                using (var controls = component.Controls)
                {
                    using (var control = controls.SingleOrDefault(item => item.Name == Model.Target.IdentifierName))
                    {
                        Debug.Assert(control != null,
                            $"input validation fail: unable to locate '{Model.Target.IdentifierName}' in Controls collection");

                        control.Name = Model.NewName;
                    }
                }
                RenameReferences(Model.Target, Model.NewName, rewriteSession);
                var controlEventHandlers = FindEventHandlersForControl(Model.Target);
                RenameDefinedFormatMembers(controlEventHandlers.ToList(), AppendUnderscoreFormat, rewriteSession);
            }
            else
            {
                RenameStandardElements(Model.Target, Model.NewName, rewriteSession);
                if (Model.Target.IsWithEvents)
                {
                    var eventHandlers = _declarationFinderProvider.DeclarationFinder.FindHandlersForWithEventsField(Model.Target);
                    RenameDefinedFormatMembers(eventHandlers.ToList(), AppendUnderscoreFormat, rewriteSession);
                }
            }
        }

        private void RenameModule(IRewriteSession rewriteSession)
        {
            RenameReferences(Model.Target, Model.NewName, rewriteSession);

            if (Model.Target.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                foreach (var reference in Model.Target.References)
                {
                    var ctxt = reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>();
                    if (ctxt != null)
                    {
                        RenameDefinedFormatMembers(_declarationFinderProvider.DeclarationFinder.FindInterfaceMembersForImplementsContext(ctxt).ToList(), AppendUnderscoreFormat, rewriteSession);
                    }
                }
            }

            var component = _projectsProvider.Component(Model.Target.QualifiedName.QualifiedModuleName);
            switch (component.Type)
            {
                case ComponentType.Document:
                    {
                        var properties = component.Properties;
                        var property = properties["_CodeName"];
                        {
                            property.Value = Model.NewName;
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
                            if ((string)property.Value == Model.Target.IdentifierName)
                            {
                                property.Value = Model.NewName;
                            }
                            component.Name = Model.NewName;
                        }
                        break;
                    }
                default:
                    {
                        using (var vbe = component.VBE)
                        {
                            if (vbe.Kind == VBEKind.Hosted)
                            {
                                // VBA - rename code module
                                using (var codeModule = component.CodeModule)
                                {
                                    Debug.Assert(!codeModule.IsWrappingNullReference,
                                        "input validation fail: Attempting to rename an ICodeModule wrapping a null reference");
                                    codeModule.Name = Model.NewName;
                                }
                            }
                            else
                            {
                                // VB6 - rename component
                                component.Name = Model.NewName;
                            }
                        }
                        break;
                    }
            }
        }

        //The parameter is not used, but it is required for the _renameActions dictionary.
        private void RenameProject(IRewriteSession rewriteSession)
        {
            var project = _projectsProvider.Project(Model.Target.ProjectId);

            if (project != null)
            {
                project.Name = Model.NewName;
            }
        }

        private void RenameDefinedFormatMembers(IReadOnlyCollection<Declaration> members, string underscoreFormat, IRewriteSession rewriteSession)
        {
            if (!members.Any()) { return; }

            var targetFragment = string.Format(underscoreFormat, Model.Target.IdentifierName);
            var replacementFragment = string.Format(underscoreFormat, Model.NewName);
            foreach (var member in members)
            {
                var newMemberName = member.IdentifierName.Replace(targetFragment, replacementFragment);
                RenameStandardElements(member, newMemberName, rewriteSession);
            }
        }

        private void RenameStandardElements(Declaration target, string newName, IRewriteSession rewriteSession)
        {
            RenameReferences(target, newName, rewriteSession);
            RenameDeclaration(target, newName, rewriteSession);
        }

        private void RenameReferences(Declaration target, string newName, IRewriteSession rewriteSession)
        {
            var modules = target.References
                .Where(reference =>
                    reference.Context.GetText() != "Me").GroupBy(r => r.QualifiedModuleName);

            foreach (var grouping in modules)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(grouping.Key);
                foreach (var reference in grouping)
                {
                    rewriter.Replace(reference.Context, newName);
                }
            }
        }

        private void RenameDeclaration(Declaration target, string newName, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedName.QualifiedModuleName);

            if (target.Context is IIdentifierContext context)
            {
                rewriter.Replace(context.IdentifierTokens, newName);
            }
        }

        private IEnumerable<Declaration> FindEventHandlersForControl(Declaration control)
        {
            if (control != null && control.DeclarationType.HasFlag(DeclarationType.Control))
            {
                return _declarationFinderProvider.DeclarationFinder.FindEventHandlers()
                    .Where(ev => ev.Scope.StartsWith($"{control.ParentScope}.{control.IdentifierName}_"));
            }
            return Enumerable.Empty<Declaration>();
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
            return _declarationFinderProvider.DeclarationFinder.FindEventHandlers()
                    .Where(ev => ev.IdentifierName.StartsWith("Class_")
                            || ev.IdentifierName.StartsWith("UserForm_")
                            || ev.IdentifierName.StartsWith("auto_"))
                    .Select(dec => dec.IdentifierName).ToList();
        }
    }
}

