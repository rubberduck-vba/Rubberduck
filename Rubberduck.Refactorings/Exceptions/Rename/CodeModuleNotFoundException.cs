using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.Rename
{
    public class CodeModuleNotFoundException : InvalidTargetDeclarationException
    {
        public CodeModuleNotFoundException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}