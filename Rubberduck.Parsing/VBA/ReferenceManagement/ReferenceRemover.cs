using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public class ReferenceRemover : ReferenceRemoverBase
    {
        private const int _maxDegreeOfReferenceRemovalParallelism = -1;

        public ReferenceRemover(
            RubberduckParserState state,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager) 
        :base(
            state,
            moduleToModuleReferenceManager)
        {}


        public override void RemoveReferencesTo(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfReferenceRemovalParallelism;

            Parallel.ForEach(modules, options, module => RemoveReferencesTo(module, token));
        }

        protected override void RemoveReferencesByFromTargetModules(IReadOnlyCollection<QualifiedModuleName> referencingModules, IReadOnlyCollection<QualifiedModuleName> targetModules, CancellationToken token)
        {
            if (!targetModules.Any())
            {
                return;
            }

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfReferenceRemovalParallelism;

            Parallel.ForEach(targetModules, options, targetModule => RemoveReferencesByFromTargetModule(referencingModules, targetModule));
        }
    }
}
