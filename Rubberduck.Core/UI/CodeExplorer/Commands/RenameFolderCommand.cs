using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.RenameFolder;
using Rubberduck.UI.CodeExplorer.Commands.Abstract;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class RenameFolderCommand : CodeExplorerInteractiveRefactoringCommandBase<RenameFolderModel>
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel)
        };

        private RubberduckParserState _state;

        public RenameFolderCommand(
            RenameFolderRefactoringAction refactoringAction,
            RefactoringUserInteraction<IRenameFolderPresenter, RenameFolderModel> userInteraction,
            RenameFolderFailedNotifier failureNotifier,
            IParserStatusProvider parserStatusProvider,
            IVbeEvents vbeEvents,
            RubberduckParserState state)
            : base(refactoringAction, userInteraction, failureNotifier, parserStatusProvider, vbeEvents)
        {
            _state = state;
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override RenameFolderModel InitialModelFromParameter(object parameter)
        {
            if (!(parameter is CodeExplorerCustomFolderViewModel folderModel))
            {
                throw new ArgumentException(nameof(parameter));
            }

            return ModelFromNode(folderModel);
        }

        private static RenameFolderModel ModelFromNode(CodeExplorerCustomFolderViewModel folderModel)
        {
            var folder = folderModel.FullPath;
            var containedModules = ContainedModules(folderModel);
            var initialSubFolder = folder.SubFolderName();
            return new RenameFolderModel(folder, containedModules, initialSubFolder);
        }

        private static ICollection<ModuleDeclaration> ContainedModules(ICodeExplorerNode itemModel)
        {
            if (itemModel is CodeExplorerComponentViewModel componentModel)
            {
                var component = componentModel.Declaration;
                return component is ModuleDeclaration moduleDeclaration
                    ? new List<ModuleDeclaration> { moduleDeclaration }
                    : new List<ModuleDeclaration>();
            }

            return itemModel.Children
                .SelectMany(ContainedModules)
                .ToList();
        }

        protected override void ValidateInitialModel(RenameFolderModel model)
        {
            var firstStaleAffectedModules = model.ModulesToMove
                .FirstOrDefault(module => _state.IsNewOrModified(module.QualifiedModuleName));
            if (firstStaleAffectedModules != null)
            {
                throw new AffectedModuleIsStaleException(firstStaleAffectedModules.QualifiedModuleName);
            }
        }

        protected override void ValidateModel(RenameFolderModel model)
        { }
    }
}