using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public interface ISupertypeClearer
    {
        void ClearSupertypes(QualifiedModuleName module);
        void ClearSupertypes(IEnumerable<QualifiedModuleName> modules);
    }
}
