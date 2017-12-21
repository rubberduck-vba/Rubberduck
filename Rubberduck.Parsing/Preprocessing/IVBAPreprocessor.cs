using Antlr4.Runtime;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface IVBAPreprocessor
    {
        void PreprocessTokenStream(QualifiedModuleName module, string moduleName, CommonTokenStream unprocessedTokenStream, BaseErrorListener errorListener, CancellationToken token);
    }
}
