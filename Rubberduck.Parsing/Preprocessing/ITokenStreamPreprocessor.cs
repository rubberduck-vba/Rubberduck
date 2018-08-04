using Antlr4.Runtime;
using System.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface ITokenStreamPreprocessor
    {
        CommonTokenStream PreprocessTokenStream(string projectId, string moduleName, CommonTokenStream tokenStream, CancellationToken token, CodeKind codeKind = CodeKind.SnippetCode);
    }
}
