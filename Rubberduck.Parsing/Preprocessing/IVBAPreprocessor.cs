using System.Threading;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface IVBAPreprocessor
    {
        string Execute(string moduleName, string unprocessedCode, CancellationToken token);
    }
}
