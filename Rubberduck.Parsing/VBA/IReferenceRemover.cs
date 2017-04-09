using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IReferenceRemover
    {
        void RemoveReferencesBy(QualifiedModuleName module, CancellationToken token);
        void RemoveReferencesBy(ICollection<QualifiedModuleName> modules, CancellationToken token);
        void RemoveReferencesTo(QualifiedModuleName module, CancellationToken token);
        void RemoveReferencesTo(ICollection<QualifiedModuleName> modules, CancellationToken token);
    }
}
