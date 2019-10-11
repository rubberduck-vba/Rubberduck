using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface ISelectedDeclarationProvider
    {
        Declaration SelectedDeclaration();
        Declaration SelectedDeclaration(QualifiedModuleName module);
        Declaration SelectedDeclaration(QualifiedSelection qualifiedSelection);
        ProjectDeclaration SelectedProject();
        ModuleDeclaration SelectedModule();
        ModuleBodyElementDeclaration SelectedMember();
    }
}