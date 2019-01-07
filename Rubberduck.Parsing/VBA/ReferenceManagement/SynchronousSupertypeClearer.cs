using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public class SynchronousSupertypeClearer : SupertypeClearerBase
    {
        public SynchronousSupertypeClearer(IDeclarationFinderProvider declarationFinderProvider) 
            :base(declarationFinderProvider)
        {}

        public override void ClearSupertypes(IEnumerable<QualifiedModuleName> modules)
        {
            foreach(var module in modules)
            {
                ClearSupertypes(module);
            }
        }
    }
}
