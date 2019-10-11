using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.VBA
{
    public class SelectedDeclarationProvider : ISelectedDeclarationProvider
    {
        private readonly ISelectionProvider _selectionProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SelectedDeclarationProvider(ISelectionProvider selectionProvider, IDeclarationFinderProvider declarationFinderProvider)
        {
            _selectionProvider = selectionProvider;
            _declarationFinderProvider = declarationFinderProvider;
        }

        public Declaration SelectedDeclaration()
        {
            var selection = _selectionProvider.ActiveSelection();
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
            var selection = _selectionProvider.Selection(module);
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
            var activeSelection = _selectionProvider.ActiveSelection();
            return activeSelection.HasValue
                ? _declarationFinderProvider.DeclarationFinder?
                    .UserDeclarations(DeclarationType.Module)
                    .OfType<ModuleDeclaration>()
                    .FirstOrDefault(module => module.QualifiedModuleName.Equals(activeSelection.Value.QualifiedName))
                : null;
        }
    }
}