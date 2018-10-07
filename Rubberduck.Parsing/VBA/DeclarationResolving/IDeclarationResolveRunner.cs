using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public interface IDeclarationResolveRunner
    {
        void CreateProjectDeclarations(IReadOnlyCollection<string> projectIds);
        void RefreshProjectReferences();
        void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
    }
}
