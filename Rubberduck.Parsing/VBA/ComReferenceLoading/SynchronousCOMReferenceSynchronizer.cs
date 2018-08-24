using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public class SynchronousCOMReferenceSynchronizer:COMReferenceSynchronizerBase 
    {
        public SynchronousCOMReferenceSynchronizer(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IProjectsProvider projectsProvider,
            IReferencedDeclarationsCollector referencedDeclarationsCollector)
        :base(
            state,
            parserStateManager,
            projectsProvider,
            referencedDeclarationsCollector)
        { }


        protected override void LoadReferences(IEnumerable<ReferenceInfo> referencesToLoad, ConcurrentBag<ReferenceInfo> unmapped, CancellationToken token)
        {
            foreach (var reference in referencesToLoad)
            {
                LoadReference(reference, unmapped);
            }
            token.ThrowIfCancellationRequested();
        }
    }
}
