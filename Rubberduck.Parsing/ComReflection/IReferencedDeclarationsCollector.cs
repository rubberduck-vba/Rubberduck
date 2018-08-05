using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IReferencedDeclarationsCollector
    {
        (IReadOnlyCollection<Declaration> declarations, Dictionary<IList<string>, Declaration>
            coClasses, SerializableProject serializableProject) CollectDeclarations(IReference reference);
    }
}
