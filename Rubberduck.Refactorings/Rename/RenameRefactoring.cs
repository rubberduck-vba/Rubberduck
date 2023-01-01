using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.Rename;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : InteractiveRefactoringBase<RenameModel>
    {
        private readonly IRefactoringAction<RenameModel> _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IProjectsProvider _projectsProvider;

        public RenameRefactoring(
            RenameRefactoringAction refactoringAction,
            RefactoringUserInteraction<IRenamePresenter, RenameModel> userInteraction,
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider, 
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _projectsProvider = projectsProvider;
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
            _refactoringAction.Refactor(model);
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

        private static RenameModel ResolveRenameTargetIfInterfaceImplementationSelected(RenameModel model)
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

