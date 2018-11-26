using System.Collections.Generic;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public class SupertypeClearer : SupertypeClearerBase
    {
        private const int _maxDegreeOfSupertypeClearingParallelism = -1;

        public SupertypeClearer(IDeclarationFinderProvider declarationFinderProvider) 
            :base(declarationFinderProvider)
        {}

        public override void ClearSupertypes(IEnumerable<QualifiedModuleName> modules)
        {
            var options = new ParallelOptions();
            options.MaxDegreeOfParallelism = _maxDegreeOfSupertypeClearingParallelism;

            Parallel.ForEach(
                   modules, 
                   options,
                   module => ClearSupertypes(module)
               );
        }
    }
}
