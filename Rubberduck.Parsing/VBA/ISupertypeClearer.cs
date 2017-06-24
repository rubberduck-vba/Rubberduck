using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    public interface ISupertypeClearer
    {
        void ClearSupertypes(QualifiedModuleName module);
        void ClearSupertypes(IEnumerable<QualifiedModuleName> modules);
    }
}
