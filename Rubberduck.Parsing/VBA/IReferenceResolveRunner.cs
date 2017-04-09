using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Threading;


namespace Rubberduck.Parsing.VBA
{
    public interface IReferenceResolveRunner
    {
        void ResolveReferences(ICollection<QualifiedModuleName> toResolve, CancellationToken token);
    }
}
