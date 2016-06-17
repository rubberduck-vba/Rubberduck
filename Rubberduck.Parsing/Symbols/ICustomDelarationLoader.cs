using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface ICustomDeclarationLoader
    {
        IReadOnlyList<Declaration> Load();
    }
}