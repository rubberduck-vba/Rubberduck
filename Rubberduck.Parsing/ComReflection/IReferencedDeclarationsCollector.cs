using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IReferencedDeclarationsCollector
    {
        IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference);
    }
}
