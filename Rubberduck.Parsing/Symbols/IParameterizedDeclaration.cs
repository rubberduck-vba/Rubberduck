using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface IParameterizedDeclaration
    {
        IEnumerable<Declaration> Parameters { get; }
        void AddParameter(Declaration parameter);
    }
}
