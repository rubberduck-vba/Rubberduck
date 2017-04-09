using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Rubberduck.Parsing.VBA
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

        public override void RemoveReferencesTo(ICollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }
            foreach(var module in modules)
            {
                RemoveReferencesTo(module, token);
            }
        }

        protected override void RemoveReferencesByFromTargetModules(ICollection<QualifiedModuleName> referencingModules, ICollection<QualifiedModuleName> targetModules, CancellationToken token)
        {
            if (!targetModules.Any())
            {
                return;
            }

            foreach(var targetModule in targetModules )
            {
                RemoveReferencesByFromTargetModule(referencingModules, targetModule);
            }
        }
    }
}
