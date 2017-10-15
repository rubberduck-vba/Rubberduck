using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    public interface IProjectReferencesProvider
    {
        IReadOnlyCollection<ReferencePriorityMap> ProjectReferences { get; }
    }
}
