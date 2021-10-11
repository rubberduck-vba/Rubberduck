using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : InteractiveRefactoringBase<ExtractInterfaceModel>
    {
        private readonly IRefactoringAction<ExtractInterfaceModel> _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ICodeBuilder _codeBuilder;
        private readonly IExtractInterfaceConflictFinderFactory _conflictFinderFactory;

        public ExtractInterfaceRefactoring(
            ExtractInterfaceRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            RefactoringUserInteraction<IExtractInterfacePresenter, ExtractInterfaceModel> userInteraction,
            ISelectionProvider selectionProvider,
            IExtractInterfaceConflictFinderFactory conflictFinderFactory,
            ICodeBuilder codeBuilder)
        :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _codeBuilder = codeBuilder;
            _conflictFinderFactory = conflictFinderFactory;
        }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.ClassModule,
            DeclarationType.Document,
            DeclarationType.UserForm
        };

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var candidates = _declarationFinderProvider.DeclarationFinder
                .Members(targetSelection.QualifiedName)
                .Where(item => ModuleTypes.Contains(item.DeclarationType));

            return candidates.SingleOrDefault(item =>
                item.QualifiedSelection.QualifiedName.Equals(targetSelection.QualifiedName));
        }

        protected override ExtractInterfaceModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ModuleTypes.Contains(target.DeclarationType) 
                || !(target is ClassModuleDeclaration targetClass))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var conflictFinder = _conflictFinderFactory.Create(_declarationFinderProvider, targetClass.ProjectId);
            var interfaceModuleName = $"I{target.IdentifierName}";

            if (conflictFinder.IsConflictingModuleName(interfaceModuleName))
            {
                interfaceModuleName = conflictFinder.GenerateNoConflictModuleName(interfaceModuleName);
            }

            var model = new ExtractInterfaceModel(_declarationFinderProvider, targetClass, _codeBuilder)
            {
                ConflictFinder = conflictFinder,
                InterfaceName = interfaceModuleName
            };

            return model;
        }

        protected override void RefactorImpl(ExtractInterfaceModel model)
        {
            _refactoringAction.Refactor(model);
        }

        //TODO: Redesign how refactoring commands are wired up to make this a responsibility of the command again. 
        public bool CanExecute(RubberduckParserState state, QualifiedModuleName qualifiedName)
        {
            var interfaceClass = state.AllUserDeclarations.SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName.Equals(qualifiedName)
                && ModuleTypes.Contains(item.DeclarationType));

            if (interfaceClass == null)
            {
                return false;
            }

            // interface class must have members to be implementable
            var hasMembers = state.AllUserDeclarations.Any(item =>
                item.DeclarationType.HasFlag(DeclarationType.Member)
                && item.Accessibility != Accessibility.Private
                && item.ParentDeclaration != null
                && item.ParentDeclaration.Equals(interfaceClass));

            if (!hasMembers)
            {
                return false;
            }

            // true if active code pane is for a class/document/form module
            return !state.IsNewOrModified(interfaceClass.QualifiedModuleName)
                   && !state.IsNewOrModified(qualifiedName);
        }
    }
}
