using System.Collections.Generic;
using System.Linq;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveToFolder;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveContainingFolderRefactoring : InteractiveRefactoringBase<IMoveMultipleFoldersPresenter, MoveMultipleFoldersModel>
    {
        private readonly IRefactoringAction<MoveMultipleFoldersModel> _moveFoldersAction;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly RubberduckParserState _state;

        public MoveContainingFolderRefactoring(
            MoveMultipleFoldersRefactoringAction moveFoldersAction,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            ISelectionProvider selectionProvider, 
            IRefactoringPresenterFactory factory, 
            IUiDispatcher uiDispatcher,
            IDeclarationFinderProvider declarationFinderProvider,
            RubberduckParserState state) 
            : base(selectionProvider, factory, uiDispatcher)
        {
            _moveFoldersAction = moveFoldersAction;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _state = state;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _selectedDeclarationProvider.SelectedModule(targetSelection);
        }

        protected override MoveMultipleFoldersModel InitializeModel(Declaration target)
        {
            if (!(target is ModuleDeclaration targetModule))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var finder = _declarationFinderProvider.DeclarationFinder;

            var sourceFolder = targetModule.CustomFolder;
            var containedModules = finder.UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .Where(module => module.ProjectId.Equals(target.ProjectId)
                                 && (module.CustomFolder.Equals(sourceFolder)
                                     || module.CustomFolder.IsSubFolderOf(sourceFolder)))
                .ToList();

            var modulesBySourceFolder = new Dictionary<string, ICollection<ModuleDeclaration>>{ {sourceFolder, containedModules} };
            var parentFolder = sourceFolder.ParentFolder();

            return new MoveMultipleFoldersModel(modulesBySourceFolder, parentFolder);
        }

        protected override void RefactorImpl(MoveMultipleFoldersModel model)
        {
            ValidateModel(model);
            _moveFoldersAction.Refactor(model);
        }

        private void ValidateModel(MoveMultipleFoldersModel model)
        {
            if (string.IsNullOrEmpty(model.TargetFolder))
            {
                throw new NoTargetFolderException();
            }

            var firstStaleAffectedModules = model.ModulesBySourceFolder.Values
                .SelectMany(modules => modules)
                .FirstOrDefault(module => _state.IsNewOrModified(module.QualifiedModuleName));
            if (firstStaleAffectedModules != null)
            {
                throw new AffectedModuleIsStaleException(firstStaleAffectedModules.QualifiedModuleName);
            }
        }
    }
}