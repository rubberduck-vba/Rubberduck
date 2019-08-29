using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.DeclarationCaching
{
    public class ConcurrentlyConstructedDeclarationFinder : DeclarationFinder
    {
        private const int _maxDegreeOfConstructionParallelism = -1;

        public ConcurrentlyConstructedDeclarationFinder(
            IReadOnlyList<Declaration> declarations, 
            IEnumerable<IAnnotation> annotations, 
            IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations,
            IReadOnlyDictionary<QualifiedModuleName, IReadOnlyCollection<IdentifierReference>> unboundDefaultMemberAccesses, 
            IHostApplication hostApp = null) 
            :base(declarations, annotations, unresolvedMemberDeclarations, unboundDefaultMemberAccesses, hostApp)
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
