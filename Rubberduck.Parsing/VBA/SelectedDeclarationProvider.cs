using System;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
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
            return FromActiveSelection(SelectedDeclaration)();
        }

        private Func<T> FromActiveSelection<T>(Func<QualifiedSelection, T> func)
            where T: class
        {
            return () =>
            {
                var activeSelection = _selectionProvider.ActiveSelection();
                return activeSelection.HasValue
                    ? func(activeSelection.Value)
                    : null;
            };
        }

        public Declaration SelectedDeclaration(QualifiedModuleName module)
        {
            return FromModuleSelection(SelectedDeclaration)(module);
        }

        private Func<QualifiedModuleName, T> FromModuleSelection<T>(Func<QualifiedSelection, T> func) 
            where T : class
        {
            return (module) =>
            {
                var selection = _selectionProvider.Selection(module);
                if (!selection.HasValue)
                {
                    return null;
                }
                var qualifiedSelection = new QualifiedSelection(module, selection.Value);
                return func(qualifiedSelection);
            };
        }

        public Declaration SelectedDeclaration(QualifiedSelection qualifiedSelection)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var canditateViaReference = SelectedDeclarationViaReference(qualifiedSelection, finder);
            if (canditateViaReference != null)
            {
                return canditateViaReference;
            }

            var canditateViaDeclaration = SelectedDeclarationViaDeclaration(qualifiedSelection, finder);
            if (canditateViaDeclaration != null)
            {
                return canditateViaDeclaration;
            }

            return SelectedModule(qualifiedSelection);
        }

        private static Declaration SelectedDeclarationViaReference(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            var referencesInModule = finder.IdentifierReferences(qualifiedSelection.QualifiedName);
            return referencesInModule
                .Where(reference => reference.IsSelected(qualifiedSelection))
                .Select(reference => reference.Declaration)
                .OrderByDescending(declaration => declaration.DeclarationType)
                // they're sorted by type, so a local comes before the procedure it's in
                .FirstOrDefault();
        }

        private static Declaration SelectedDeclarationViaDeclaration(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            var declarationsInModule = finder.Members(qualifiedSelection.QualifiedName);
            return declarationsInModule
                .Where(declaration => declaration.IsSelected(qualifiedSelection))
                .OrderByDescending(declaration => declaration.DeclarationType)
                // they're sorted by type, so a local comes before the procedure it's in
                .FirstOrDefault();
        }

        public ModuleBodyElementDeclaration SelectedMember()
        {
            return FromActiveSelection(SelectedMember)();
        }

        public ModuleBodyElementDeclaration SelectedMember(QualifiedModuleName module)
        {
            return FromModuleSelection(SelectedMember)(module);
        }

        public ModuleBodyElementDeclaration SelectedMember(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Member)
                .OfType<ModuleBodyElementDeclaration>()
                .FirstOrDefault(member => member.QualifiedModuleName.Equals(qualifiedSelection.QualifiedName)
                                          && member.Context.GetSelection().Contains(qualifiedSelection.Selection));
        }

        public ModuleDeclaration SelectedModule()
        {
            return FromActiveSelection(SelectedModule)();
        }

        public ModuleDeclaration SelectedModule(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .FirstOrDefault(module => module.QualifiedModuleName.Equals(qualifiedSelection.QualifiedName));
        }

        public ProjectDeclaration SelectedProject()
        {
            return FromActiveSelection(SelectedProject)();
        }

        public ProjectDeclaration SelectedProject(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Project)
                .OfType<ProjectDeclaration>()
                .FirstOrDefault(project => project.ProjectId.Equals(qualifiedSelection.QualifiedName.ProjectId));
        }
    }
}