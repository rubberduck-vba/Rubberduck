using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousCOMReferenceManager:COMReferenceManagerBase 
    {
        public SynchronousCOMReferenceManager(RubberduckParserState state, IParserStateManager parserStateManager, string serializedDeclarationsPath = null)
        :base(state, parserStateManager, serializedDeclarationsPath) { }


        protected override void LoadReferences(List<IReference> referencesToLoad, ConcurrentBag<IReference> unmapped, CancellationToken token)
        {
            foreach (var reference in referencesToLoad)
            {
                var localReference = reference;
                LoadReference(localReference, unmapped);
            }
            token.ThrowIfCancellationRequested();
        }
    }
}
