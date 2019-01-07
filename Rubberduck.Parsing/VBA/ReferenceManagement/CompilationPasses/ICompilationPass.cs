using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement.CompilationPasses
{
    public interface ICompilationPass
    {
        void Execute(IReadOnlyCollection<QualifiedModuleName> modules);
    }
}
