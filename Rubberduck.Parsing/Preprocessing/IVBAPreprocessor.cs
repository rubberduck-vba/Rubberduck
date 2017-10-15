using Antlr4.Runtime;
using System.Threading;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface IVBAPreprocessor
    {
        void PreprocessTokenStream(string moduleName, CommonTokenStream unprocessedTokenStream, CancellationToken token);
    }
}
