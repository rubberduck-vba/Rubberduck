using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public interface IParseRunner
    {
        void ParseModules(IReadOnlyCollection<QualifiedModuleName> modulesToParse, CancellationToken token);
    }
}
