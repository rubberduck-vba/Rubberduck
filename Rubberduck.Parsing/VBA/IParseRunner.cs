using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IParseRunner
    {
        void ParseModules(ICollection<QualifiedModuleName> modulesToParse, CancellationToken token);
    }
}
