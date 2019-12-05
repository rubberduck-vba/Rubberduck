using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.Rename;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using NLog;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : InteractiveRefactoringBase<IRenamePresenter, RenameModel>
    {
        private const string AppendUnderscoreFormat = "{0}_";
        private const string PrependUnderscoreFormat = "_{0}";

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IDictionary<DeclarationType, Action<RenameModel, IRewriteSession>> _renameActions;

        private readonly IParseManager _parseManager;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RenameRefactoring(
            IRefactoringPresenterFactory factory, 
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider, 
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IParseManager parseManager,
            IUiDispatcher uiDispatcher)
        :base(rewritingManager, selectionProvider, factory, uiDispatcher)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _projectsProvider = projectsProvider;
            _parseManager = parseManager;

            _renameActions = new Dictionary<DeclarationType, Action<RenameModel, IRewriteSession>>
            {
                {DeclarationType.Member, RenameMember},
                {DeclarationType.Parameter, RenameParameter},
                {DeclarationType.Event, RenameEvent},
                {DeclarationType.Variable, RenameVariable},
                {DeclarationType.Module, RenameModule},
                {DeclarationType.Project, RenameProject}
            };
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
        }

        protected override RenameModel InitializeModel(Declaration target)
        {
            CheckWhetherValidTarget(target);

            var model = DeriveTarget(new RenameModel(target));

            if (!model.Target.Equals(model.InitialTarget))
            {
                CheckWhetherValidTarget(model.Target);
            }

            return model;
        }

        protected override void RefactorImpl(RenameModel model)
        {
            if (model.Target.DeclarationType.HasFlag(DeclarationType.Module)
                || model.Target.DeclarationType.HasFlag(DeclarationType.Project))
            {
                //The parser needs to be suspended during the refactoring of a component because the VBE API object rename causes a separate reparse. 
                RenameRefactorWithSuspendedParser(model);
                return;
            }

            RenameRefactor(model);
        }

        private void RenameRefactorWithSuspendedParser(RenameModel model)
        {
            var suspendResult = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, () => RenameRefactor(model));
            var suspendOutcome = suspendResult.Outcome;
            if (suspendOutcome != SuspensionOutcome.Completed)
            {
                if ((suspendOutcome == SuspensionOutcome.UnexpectedError || suspendOutcome == SuspensionOutcome.Canceled)
                    && suspendResult.EncounteredException != null)
                {
                    ExceptionDispatchInfo.Capture(suspendResult.EncounteredException).Throw();
                    return;
                }

                _logger.Warn($"{nameof(RenameRefactor)} failed because a parser suspension request could not be fulfilled.  The request's result was '{suspendResult.ToString()}'.");
                throw new SuspendParserFailureException();
            }
        }

        private void RenameRefactor(RenameModel model)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            Rename(model, rewriteSession);
            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private RenameModel DeriveTarget(RenameModel model)
        {
            if (!(model.InitialTarget is IInterfaceExposable))
            {
                return model;
            }

            return ResolveRenameTargetIfEventHandlerSelected(model) 
                   ?? ResolveRenameTargetIfInterfaceImplementationSelected(model) 
                   ?? model;
        }

        private RenameModel ResolveRenameTargetIfEventHandlerSelected(RenameModel model)
        {
            var initialTarget = model.InitialTarget;

            if (initialTarget.DeclarationType.HasFlag(DeclarationType.Procedure) 
                && initialTarget.IdentifierName.Contains("_"))
            {
                return ResolveEventHandlerToControl(model) ??
                       ResolveEventHandlerToUserEvent(model);
            }
            return null;
        }

        private void CheckWhetherValidTarget(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!target.IsUserDefined)
            {
                throw new TargetDeclarationNotUserDefinedException(target);
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
                            throw new TargetControlNotFoundException(target);
                        }
                    }
                }
            }

            if (target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                var component = _projectsProvider.Component(target.QualifiedName.QualifiedModuleName);
                using (var module = component.CodeModule)
                {
                    if (module.IsWrappingNullReference)
                    {
                        throw new CodeModuleNotFoundException(target);
                    }
                }
            }

            if (target is IInterfaceExposable && StandardEventHandlerNames().Contains(target.IdentifierName))
            {
                throw new TargetDeclarationIsStandardEventHandlerException(target);
            }
        }

        private RenameModel ResolveEventHandlerToControl(RenameModel model)
        {
            var userTarget = model.InitialTarget;
            var control = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Control)
                .FirstOrDefault(ctrl => userTarget.Scope.StartsWith($"{ctrl.ParentScope}.{ctrl.IdentifierName}_"));

            if (FindEventHandlersForControl(control).Contains(userTarget))
            {
                model.IsControlEventHandlerRename = true;
                model.Target = control;
                return model;
            }

            return null;
        }

        private RenameModel ResolveEventHandlerToUserEvent(RenameModel model)
        {
            var userTarget = model.InitialTarget;

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

                        model.IsUserEventHandlerRename = eventDeclaration != null;

                        if (eventDeclaration != null)
                        {
                            model.Target = eventDeclaration;
                            return model;
                        }
                    }
                }
            }
            return null;
        }

        private RenameModel ResolveRenameTargetIfInterfaceImplementationSelected(RenameModel model)
        {
            var userTarget = model.InitialTarget;

            var interfaceMember = userTarget is IInterfaceExposable member && member.IsInterfaceMember
                ? userTarget
                : (userTarget as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

            if (interfaceMember != null)
            {
                model.IsInterfaceMemberRename = true;
                model.Target = interfaceMember;
                return model;
            }

            return null;
        }

        private void Rename(RenameModel model, IRewriteSession rewriteSession)
        {
            Debug.Assert(!model.NewName.Equals(model.Target.IdentifierName, StringComparison.InvariantCultureIgnoreCase),
                            $"input validation fail: New Name equals Original Name ({model.Target.IdentifierName})");

            var actionKeys = _renameActions.Keys.Where(decType => model.Target.DeclarationType.HasFlag(decType)).ToList();
            if (actionKeys.Any())
            {
                Debug.Assert(actionKeys.Count == 1, $"{actionKeys.Count} Rename Actions have flag '{model.Target.DeclarationType.ToString()}'");
                _renameActions[actionKeys.FirstOrDefault()](model, rewriteSession);
            }
            else
            {
                RenameStandardElements(model.Target, model.NewName, rewriteSession);
            }
        }

        private void RenameMember(RenameModel model, IRewriteSession rewriteSession)
        {
            if (model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var members = _declarationFinderProvider.DeclarationFinder.MatchName(model.Target.IdentifierName)
                    .Where(item => item.ProjectId == model.Target.ProjectId
                        && item.ComponentName == model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    RenameStandardElements(member, model.NewName, rewriteSession);
                }
            }
            else
            {
                RenameStandardElements(model.Target, model.NewName, rewriteSession);
            }

            if (!model.IsInterfaceMemberRename)
            {
                return;
            }
            
            var implementations = _declarationFinderProvider.DeclarationFinder.FindAllInterfaceImplementingMembers()
                .Where(impl => ReferenceEquals(model.Target.ParentDeclaration, impl.InterfaceImplemented)
                               && impl.InterfaceMemberImplemented.IdentifierName.Equals(model.Target.IdentifierName));

            RenameDefinedFormatMembers(model, implementations.ToList(), PrependUnderscoreFormat, rewriteSession);
        }

        private void RenameParameter(RenameModel model, IRewriteSession rewriteSession)
        {
            if (model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var parameters = _declarationFinderProvider.DeclarationFinder.MatchName(model.Target.IdentifierName).Where(param =>
                   param.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                   && param.DeclarationType == DeclarationType.Parameter
                    && param.ParentDeclaration.IdentifierName.Equals(model.Target.ParentDeclaration.IdentifierName)
                    && param.ParentDeclaration.ParentScopeDeclaration.Equals(model.Target.ParentDeclaration.ParentScopeDeclaration));

                foreach (var param in parameters)
                {
                    RenameStandardElements(param, model.NewName, rewriteSession);
                }
            }
            else
            {
                RenameStandardElements(model.Target, model.NewName, rewriteSession);
            }
        }

        private void RenameEvent(RenameModel model, IRewriteSession rewriteSession)
        {
            RenameStandardElements(model.Target, model.NewName, rewriteSession);

            var withEventsDeclarations = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(varDec => varDec.IsWithEvents && varDec.AsTypeName.Equals(model.Target.ParentDeclaration.IdentifierName));

            var eventHandlers = withEventsDeclarations.SelectMany(we => _declarationFinderProvider.DeclarationFinder.FindHandlersForWithEventsField(we));
            RenameDefinedFormatMembers(model, eventHandlers.ToList(), PrependUnderscoreFormat, rewriteSession);
        }

        private void RenameVariable(RenameModel model, IRewriteSession rewriteSession)
        {
            if ((model.Target.Accessibility == Accessibility.Public ||
                 model.Target.Accessibility == Accessibility.Implicit)
                && model.Target.ParentDeclaration is ClassModuleDeclaration classDeclaration
                && classDeclaration.Subtypes.Any())
            {
                RenameMember(model, rewriteSession);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Control))
            {
                var component = _projectsProvider.Component(model.Target.QualifiedName.QualifiedModuleName);
                using (var controls = component.Controls)
                {
                    using (var control = controls.SingleOrDefault(item => item.Name == model.Target.IdentifierName))
                    {
                        Debug.Assert(control != null,
                            $"input validation fail: unable to locate '{model.Target.IdentifierName}' in Controls collection");

                        control.Name = model.NewName;
                    }
                }
                RenameReferences(model.Target, model.NewName, rewriteSession);
                var controlEventHandlers = FindEventHandlersForControl(model.Target);
                RenameDefinedFormatMembers(model, controlEventHandlers.ToList(), AppendUnderscoreFormat, rewriteSession);
            }
            else
            {
                RenameStandardElements(model.Target, model.NewName, rewriteSession);
                if (model.Target.IsWithEvents)
                {
                    var eventHandlers = _declarationFinderProvider.DeclarationFinder.FindHandlersForWithEventsField(model.Target);
                    RenameDefinedFormatMembers(model, eventHandlers.ToList(), AppendUnderscoreFormat, rewriteSession);
                }
            }
        }

        private void RenameModule(RenameModel model, IRewriteSession rewriteSession)
        {
            RenameReferences(model.Target, model.NewName, rewriteSession);

            if (model.Target.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                foreach (var reference in model.Target.References)
                {
                    var ctxt = reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>();
                    if (ctxt != null)
                    {
                        RenameDefinedFormatMembers(model, _declarationFinderProvider.DeclarationFinder.FindInterfaceMembersForImplementsContext(ctxt).ToList(), AppendUnderscoreFormat, rewriteSession);
                    }
                }
            }

            var component = _projectsProvider.Component(model.Target.QualifiedName.QualifiedModuleName);
            switch (component.Type)
            {
                case ComponentType.Document:
                    {
                        using (var properties = component.Properties)
                        using (var property = properties["_CodeName"])
                        {
                            property.Value = model.NewName;
                        }
                        break;
                    }
                case ComponentType.UserForm:
                case ComponentType.VBForm:
                case ComponentType.MDIForm:
                    {
                        using (var properties = component.Properties)
                        using (var property = properties["Caption"])
                        {
                            if ((string)property.Value == model.Target.IdentifierName)
                            {
                                property.Value = model.NewName;
                            }
                            component.Name = model.NewName;
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
                                    codeModule.Name = model.NewName;
                                }
                            }
                            else
                            {
                                // VB6 - rename component
                                component.Name = model.NewName;
                            }
                        }
                        break;
                    }
            }
        }

        //TODO: Implement renaming references to the project in code.
        private void RenameProject(RenameModel model, IRewriteSession rewriteSession)
        {
            var project = _projectsProvider.Project(model.Target.ProjectId);

            if (project != null)
            {
                project.Name = model.NewName;
            }
        }

        private void RenameDefinedFormatMembers(RenameModel model, IReadOnlyCollection<Declaration> members, string underscoreFormat, IRewriteSession rewriteSession)
        {
            if (!members.Any()) { return; }

            var targetFragment = string.Format(underscoreFormat, model.Target.IdentifierName);
            var replacementFragment = string.Format(underscoreFormat, model.NewName);
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
                    reference.Context.GetText() != "Me" 
                    && !reference.IsArrayAccess
                    && !reference.IsDefaultMemberAccess)
                .GroupBy(r => r.QualifiedModuleName);

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

        private List<string> StandardEventHandlerNames()
        {
            return _declarationFinderProvider.DeclarationFinder.FindEventHandlers()
                    .Where(ev => ev.IdentifierName.StartsWith("Class_")
                            || ev.IdentifierName.StartsWith("UserForm_")
                            || ev.IdentifierName.StartsWith("auto_"))
                    .Select(dec => dec.IdentifierName).ToList();
        }
    }
}

