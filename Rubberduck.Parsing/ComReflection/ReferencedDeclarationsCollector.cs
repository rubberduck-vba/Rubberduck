using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public abstract class ReferencedDeclarationsCollectorBase : IReferencedDeclarationsCollector
    {
        private readonly IDeclarationsFromComProjectLoader _declarationsFromComProjectLoader;

        protected ReferencedDeclarationsCollectorBase(IDeclarationsFromComProjectLoader declarationsFromComProjectLoader)
        {
            _declarationsFromComProjectLoader = declarationsFromComProjectLoader;
        }


        public abstract IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference);


        protected IReadOnlyCollection<Declaration> LoadDeclarationsFromComProject(ComProject type, string projectId = null)
        {
            return _declarationsFromComProjectLoader.LoadDeclarations(type, projectId);
        } 
    }
}
