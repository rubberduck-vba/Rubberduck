using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public interface IStringParser
    {
        (IParseTree tree, ITokenStream tokenStream) Parse(string moduleName, string projectId, string code, CancellationToken token, CodeKind codeKind = CodeKind.SnippetCode, ParserMode parserMode = ParserMode.FallBackSllToLl);
    }
}
