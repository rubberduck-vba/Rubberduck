using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface IParameterizedDeclaration
    {
        IReadOnlyList<ParameterDeclaration> Parameters { get; }
        void AddParameter(ParameterDeclaration parameter);
    }
}
