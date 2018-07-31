using Antlr4.Runtime;
using System.Threading;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface IVBAPreprocessor
    {
        void PreprocessTokenStream(string projectId, string moduleName, CommonTokenStream unprocessedTokenStream, BaseErrorListener errorListener, CancellationToken token);
    }
}
