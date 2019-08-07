using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.Rename
{
    public class TargetControlNotFoundException : InvalidTargetDeclarationException
    {
        public TargetControlNotFoundException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}
