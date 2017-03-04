using System.Threading;

namespace Rubberduck.Parsing.Preprocessing
{
    public interface IVBAPreprocessor
    {
        string Execute(string moduleName, string unprocessedCode, CancellationToken token);
    }
}
