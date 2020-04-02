using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : InteractiveRefactoringBase<ReorderParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IRefactoringAction<ReorderParametersModel> _refactoringAction;

        public ReorderParametersRefactoring(
            ReorderParameterRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            RefactoringUserInteraction<IReorderParametersPresenter, ReorderParametersModel> userInteraction,
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

        protected override ReorderParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var model = DerivedTarget(new ReorderParametersModel(target));

            return model;
        }

        private ReorderParametersModel DerivedTarget(ReorderParametersModel model)
        {
            var preliminaryModel = ResolvedInterfaceMemberTarget(model)
                                   ?? ResolvedEventTarget(model)
                                   ?? model;
            return ResolvedGetterTarget(preliminaryModel) ?? preliminaryModel;
        }

        private static ReorderParametersModel ResolvedInterfaceMemberTarget(ReorderParametersModel model)
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

        private ReorderParametersModel ResolvedEventTarget(ReorderParametersModel model)
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

        private ReorderParametersModel ResolvedGetterTarget(ReorderParametersModel model)
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

        protected override void RefactorImpl(ReorderParametersModel model)
        {
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
