using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.Rename;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : InteractiveRefactoringBase<IRenamePresenter, RenameModel>
    {
        private const string AppendUnderscoreFormat = "{0}_";
        private const string PrependUnderscoreFormat = "_{0}";

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IDictionary<DeclarationType, Action<IRewriteSession>> _renameActions;
        private readonly List<string> _standardEventHandlerNames;


        public RenameRefactoring(IRefactoringPresenterFactory factory, IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _projectsProvider = projectsProvider;

            _renameActions = new Dictionary<DeclarationType, Action<IRewriteSession>>
            {
                {DeclarationType.Member, RenameMember},
                {DeclarationType.Parameter, RenameParameter},
                {DeclarationType.Event, RenameEvent},
                {DeclarationType.Variable, RenameVariable},
                {DeclarationType.Module, RenameModule},
                {DeclarationType.Project, RenameProject}
            };

            _standardEventHandlerNames = _declarationFinderProvider != null
                ? StandardEventHandlerNames()
                : new List<string>();
        }

        public override void Refactor(QualifiedSelection targetSelection)
        {
            var target = _declarationFinderProvider.DeclarationFinder
                .FindSelectedDeclaration(targetSelection);

            if (target == null)
            {
                throw new NoDeclarationForSelectionException(targetSelection);
            }

            Refactor(target);
        }

        public override void Refactor(Declaration target)
        {
            Refactor(InitializeModel(target));
        }

        private RenameModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException(target);
            }

            CheckWhetherValidTarget(target);

            var model = DeriveTarget(new RenameModel(target));

            if (!model.Target.Equals(model.InitialTarget))
            {
                CheckWhetherValidTarget(model.Target);
            }

            return model;
        }

        protected override void RefactorImpl(IRenamePresenter presenter)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            Rename(rewriteSession);
            rewriteSession.TryRewrite();
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
                throw new InvalidTargetDeclarationException(null);
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

            if (target is IInterfaceExposable && _standardEventHandlerNames.Contains(target.IdentifierName))
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

            if (!Model.IsInterfaceMemberRename)
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

