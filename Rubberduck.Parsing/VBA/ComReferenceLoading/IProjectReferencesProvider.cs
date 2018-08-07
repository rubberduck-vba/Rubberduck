using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public interface IProjectReferencesProvider
    {
        IReadOnlyCollection<ReferencePriorityMap> ProjectReferences { get; }
    }
}
