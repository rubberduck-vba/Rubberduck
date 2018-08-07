using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
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
