using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousCOMReferenceSynchronizer:COMReferenceSynchronizerBase 
    {
        public SynchronousCOMReferenceSynchronizer(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            string serializedDeclarationsPath = null)
        :base(
            state,
            parserStateManager,
            serializedDeclarationsPath)
        { }


        protected override void LoadReferences(IEnumerable<IReference> referencesToLoad, ConcurrentBag<IReference> unmapped, CancellationToken token)
        {
            foreach (var reference in referencesToLoad)
            {
                LoadReference(reference, unmapped);
            }
            token.ThrowIfCancellationRequested();
        }
    }
}
