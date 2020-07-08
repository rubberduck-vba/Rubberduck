using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.RemoveParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersRefactoring : InteractiveRefactoringBase<RemoveParametersModel>
    {
        private readonly IRefactoringAction<RemoveParametersModel> _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public RemoveParametersRefactoring(
            RemoveParameterRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            RefactoringUserInteraction<IRemoveParametersPresenter, RemoveParametersModel> userInteraction,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
        :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (!ValidDeclarationTypes.Contains(selectedDeclaration.DeclarationType))
            {
                return selectedDeclaration.DeclarationType == DeclarationType.Parameter
                    ? _selectedDeclarationProvider.SelectedMember(targetSelection)
                    : null;
            }

            return selectedDeclaration;
        }

        protected override RemoveParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ValidDeclarationTypes.Contains(target.DeclarationType) && target.DeclarationType != DeclarationType.Parameter)
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var model = DerivedTarget(new RemoveParametersModel(target));

            return model;
        }

        private RemoveParametersModel DerivedTarget(RemoveParametersModel model)
        {
            var preliminaryModel = ResolvedInterfaceMemberTarget(model) 
                                   ?? ResolvedEventTarget(model) 
                                   ?? model;
            return ResolvedGetterTarget(preliminaryModel) ?? preliminaryModel;
        }

        private static RemoveParametersModel ResolvedInterfaceMemberTarget(RemoveParametersModel model)
        {
            var declaration = model.TargetDeclaration;
            if (!(declaration is ModuleBodyElementDeclaration member) || !member.IsInterfaceImplementation)
            {
                return null;
            }

            model.IsInterfaceMemberRefactoring = true;
            model.TargetDeclaration = member.InterfaceMemberImplemented;

            return model;
        }

        private RemoveParametersModel ResolvedEventTarget(RemoveParametersModel model)
        {
            foreach (var eventDeclaration in _declarationFinderProvider
                .DeclarationFinder
                .UserDeclarations(DeclarationType.Event))
            {
                if (_declarationFinderProvider.DeclarationFinder
                    .FindEventHandlers(eventDeclaration)
                    .Any(handler => Equals(handler, model.TargetDeclaration)))
                {
                    model.IsEventRefactoring = true;
                    model.TargetDeclaration = eventDeclaration;
                    return model;
                }
            }
            return null;
        }

        private RemoveParametersModel ResolvedGetterTarget(RemoveParametersModel model)
        {
            var target = model.TargetDeclaration;
            if (target == null || !target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                return null;
            }

            if (target.DeclarationType == DeclarationType.PropertyGet)
            {
                model.IsPropertyRefactoringWithGetter = true;
                return model;
            }


            var getter = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.PropertyGet)
                .FirstOrDefault(item => item.Scope == target.Scope 
                                        && item.IdentifierName == target.IdentifierName);

            if (getter == null)
            {
                return null;
            }

            model.IsPropertyRefactoringWithGetter = true;
            model.TargetDeclaration = getter;

            return model;
        }

        protected override void RefactorImpl(RemoveParametersModel model)
        {
            if (model.TargetDeclaration == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            _refactoringAction.Refactor(model);
        }

        public void QuickFix(QualifiedSelection selection)
        {
            var targetDeclaration = FindTargetDeclaration(selection);
            var model = InitializeModel(targetDeclaration);
            
            var selectedParameters = model.Parameters.Where(p => selection.Selection.Contains(p.Declaration.QualifiedSelection.Selection)).ToList();

            if (selectedParameters.Count > 1)
            {
                throw new MultipleParametersSelectedException(selectedParameters);
            }

            var target = selectedParameters.SingleOrDefault(p => selection.Selection.Contains(p.Declaration.QualifiedSelection.Selection));

            if (target == null)
            {
                throw new NoParameterSelectedException();
            }

            model.RemoveParameters.Add(target);
            _refactoringAction.Refactor(model);
        }

        public static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
    }
}
