using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IDeclarationsFromComProjectLoader
    {
        IReadOnlyCollection<Declaration> LoadDeclarations(ComProject type, string projectId = null);
    }
}