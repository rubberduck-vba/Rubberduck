using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface IDeclarationWithParameter
    {
        IEnumerable<Declaration> Parameters { get; }
        void Add(Declaration parameter);
    }
}
