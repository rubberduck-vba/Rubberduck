using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.VBA
{
    public class SelectedDeclarationService : ISelectedDeclarationService
    {
        private readonly ISelectionService _selectionService;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SelectedDeclarationService(ISelectionService selectionService, IDeclarationFinderProvider declarationFinderProvider)
        {
            _selectionService = selectionService;
            _declarationFinderProvider = declarationFinderProvider;
        }

        public Declaration SelectedDeclaration()
        {
            var selection = _selectionService.ActiveSelection();
            return SelectedDeclaration(selection);
        }

        private Declaration SelectedDeclaration(QualifiedSelection? qualifiedSelection)
        {
            return qualifiedSelection.HasValue
                ? SelectedDeclaration(qualifiedSelection.Value)
                : null;
        }

        public Declaration SelectedDeclaration(QualifiedModuleName module)
        {
            var selection = _selectionService.Selection(module);
            return SelectedDeclaration(module, selection);
        }

        private Declaration SelectedDeclaration(QualifiedModuleName module, Selection? selection)
        {
            return selection.HasValue
                ? SelectedDeclaration(new QualifiedSelection(module, selection.Value))
                : null;
        }

        public Declaration SelectedDeclaration(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?.FindSelectedDeclaration(qualifiedSelection);
        }

        public ModuleDeclaration SelectedModule()
        {
            var activeSelection = _selectionService.ActiveSelection();
            return activeSelection.HasValue
                ? _declarationFinderProvider.DeclarationFinder?
                    .UserDeclarations(DeclarationType.Module)
                    .OfType<ModuleDeclaration>()
                    .FirstOrDefault(module => module.QualifiedModuleName.Equals(activeSelection.Value.QualifiedName))
                : null;
        }
    }
}