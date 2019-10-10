using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface ISelectedDeclarationService
    {
        Declaration SelectedDeclaration();
        Declaration SelectedDeclaration(QualifiedModuleName module);
        Declaration SelectedDeclaration(QualifiedSelection qualifiedSelection);
        ModuleDeclaration SelectedModule();
    }
}