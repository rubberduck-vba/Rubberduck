using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface ISelectedDeclarationProvider
    {
        Declaration SelectedDeclaration();
        Declaration SelectedDeclaration(QualifiedModuleName module);
        Declaration SelectedDeclaration(QualifiedSelection qualifiedSelection);
        ModuleBodyElementDeclaration SelectedMember();
        ModuleBodyElementDeclaration SelectedMember(QualifiedModuleName module);
        ModuleBodyElementDeclaration SelectedMember(QualifiedSelection qualifiedSelection);
        ProjectDeclaration SelectedProject();
        ProjectDeclaration SelectedProject(QualifiedSelection qualifiedSelection);
        ModuleDeclaration SelectedModule();
        ModuleDeclaration SelectedModule(QualifiedSelection qualifiedSelection);

        ModuleDeclaration SelectedProjectExplorerModule();
    }
}