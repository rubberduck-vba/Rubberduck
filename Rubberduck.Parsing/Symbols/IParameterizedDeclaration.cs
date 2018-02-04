using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface IParameterizedDeclaration
    {
        IEnumerable<ParameterDeclaration> Parameters { get; }
        void AddParameter(ParameterDeclaration parameter);
    }
}
