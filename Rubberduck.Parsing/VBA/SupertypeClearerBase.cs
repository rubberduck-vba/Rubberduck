using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public abstract class SupertypeClearerBase : ISupertypeClearer
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        protected SupertypeClearerBase(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider ?? throw new ArgumentNullException(nameof(declarationFinderProvider));
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
