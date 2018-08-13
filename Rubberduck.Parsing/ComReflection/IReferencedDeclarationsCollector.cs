using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IReferencedDeclarationsCollector
    {
        IReadOnlyCollection<Declaration> CollectedDeclarations(IReference reference);
    }
}
