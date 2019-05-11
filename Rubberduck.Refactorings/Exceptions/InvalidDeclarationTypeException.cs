using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions
{
    public class InvalidDeclarationTypeException : InvalidTargetDeclarationException
    {
        public InvalidDeclarationTypeException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}