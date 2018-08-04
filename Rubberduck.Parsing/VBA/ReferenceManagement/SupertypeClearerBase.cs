using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public abstract class SupertypeClearerBase : ISupertypeClearer
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SupertypeClearerBase(IDeclarationFinderProvider declarationFinderProvider)
        {
            if(declarationFinderProvider == null)
            {
                throw new ArgumentNullException(nameof(declarationFinderProvider));
            }
            _declarationFinderProvider = declarationFinderProvider;
        }

        public abstract void ClearSupertypes(IEnumerable<QualifiedModuleName> modules);

        public void ClearSupertypes(QualifiedModuleName module)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var moduleDeclaration = finder.ModuleDeclaration(module);
            (moduleDeclaration as ClassModuleDeclaration)?.ClearSupertypes();   
        }
    }
}
