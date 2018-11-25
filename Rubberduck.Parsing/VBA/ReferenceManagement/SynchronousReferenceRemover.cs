using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public class SynchronousReferenceRemover : ReferenceRemoverBase 
    {
        public SynchronousReferenceRemover(
            RubberduckParserState state,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager)
        :base(
            state,
            moduleToModuleReferenceManager)
        { }


        public override void RemoveReferencesTo(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            foreach(var module in modules)
            {
                RemoveReferencesTo(module, token);
            }
        }

        protected override void RemoveReferencesByFromTargetModules(IReadOnlyCollection<QualifiedModuleName> referencingModules, IReadOnlyCollection<QualifiedModuleName> targetModules, CancellationToken token)
        {
            foreach(var targetModule in targetModules )
            {
                RemoveReferencesByFromTargetModule(referencingModules, targetModule);
            }
        }
    }
}
