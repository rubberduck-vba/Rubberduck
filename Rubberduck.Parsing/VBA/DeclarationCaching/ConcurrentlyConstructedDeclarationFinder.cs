using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.DeclarationCaching
{
    public class ConcurrentlyConstructedDeclarationFinder : DeclarationFinder
    {
        private const int _maxDegreeOfConstructionParallelism = -1;
        
        public ConcurrentlyConstructedDeclarationFinder(IReadOnlyList<Declaration> declarations,
            IEnumerable<IParseTreeAnnotation> annotations,
            IReadOnlyDictionary<QualifiedModuleName, LogicalLineStore> logicalLines,
            IReadOnlyDictionary<QualifiedModuleName, IFailedResolutionStore> failedResolutionStores,
            IHostApplication hostApp = null) 
            :base(declarations, annotations, logicalLines, failedResolutionStores, hostApp)
        {}

        protected override void ExecuteCollectionConstructionActions(List<Action> collectionConstructionActions)
        {
            var options = new ParallelOptions();
            options.MaxDegreeOfParallelism = _maxDegreeOfConstructionParallelism;

            Parallel.ForEach(
                collectionConstructionActions, 
                options,
                action => action.Invoke() 
            );
        }
    }
}
