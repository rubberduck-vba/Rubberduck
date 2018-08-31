using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface ICompilationPass
    {
        void Execute(IReadOnlyCollection<QualifiedModuleName> modules);
    }
}
