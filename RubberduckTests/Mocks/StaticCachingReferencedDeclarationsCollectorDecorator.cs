using System.Collections.Generic;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace RubberduckTests.Mocks
{
    public class StaticCachingReferencedDeclarationsCollectorDecorator : IReferencedDeclarationsCollector
    {
        private static readonly Dictionary<ReferenceInfo, IReadOnlyCollection<Declaration>> CachedReferences = new Dictionary<ReferenceInfo, IReadOnlyCollection<Declaration>>();

        private IReferencedDeclarationsCollector _baseCollector;

        public StaticCachingReferencedDeclarationsCollectorDecorator(IReferencedDeclarationsCollector baseCollector)
        {
            _baseCollector = baseCollector;
        }

        public IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference)
        {
            if (CachedReferences.TryGetValue(reference, out var cachedDeclarations))
            {
                foreach (var declaration in cachedDeclarations)
                {
                    declaration.ClearReferences();
                }
                return cachedDeclarations;
            }

            var collectedDeclarations = _baseCollector.CollectedDeclarations(reference);
            CachedReferences[reference] = collectedDeclarations;
            return collectedDeclarations;
        }
    }
}