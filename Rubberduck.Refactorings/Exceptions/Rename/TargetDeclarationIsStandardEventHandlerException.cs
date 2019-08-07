using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.Rename
{
    public class TargetDeclarationIsStandardEventHandlerException : InvalidTargetDeclarationException
    {
        public TargetDeclarationIsStandardEventHandlerException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}