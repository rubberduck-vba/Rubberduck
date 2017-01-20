using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
{
    public interface ICustomDeclarationLoader
    {
        IReadOnlyList<Declaration> Load();
    }
}