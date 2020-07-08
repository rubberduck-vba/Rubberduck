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

        public ExtractInterfaceRefactoring(
            ExtractInterfaceRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            RefactoringUserInteraction<IExtractInterfacePresenter, ExtractInterfaceModel> userInteraction,
            ISelectionProvider selectionProvider,
            ICodeBuilder codeBuilder)
        :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _codeBuilder = codeBuilder;
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

            return new ExtractInterfaceModel(_declarationFinderProvider, targetClass, _codeBuilder);
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
                && item.ParentDeclaration != null
                && item.ParentDeclaration.Equals(interfaceClass));

            if (!hasMembers)
            {
                return false;
            }

            var parseTree = state.GetParseTree(interfaceClass.QualifiedName.QualifiedModuleName);
            var context = ((Antlr4.Runtime.ParserRuleContext)parseTree).GetDescendents<VBAParser.ImplementsStmtContext>();

            // true if active code pane is for a class/document/form module
            return !context.Any()
                   && !state.IsNewOrModified(interfaceClass.QualifiedModuleName)
                   && !state.IsNewOrModified(qualifiedName);
        }
    }
}
