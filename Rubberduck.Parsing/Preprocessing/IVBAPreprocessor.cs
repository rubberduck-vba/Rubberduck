using Antlr4.Runtime;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Threading;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface IVBAPreprocessor
    {
        void PreprocessTokenStream(IVBProject project, string moduleName, CommonTokenStream unprocessedTokenStream, BaseErrorListener errorListener, CancellationToken token);
    }
}
