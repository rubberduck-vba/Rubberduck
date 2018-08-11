using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public interface ITokenStreamParser
    {
        IParseTree Parse(string moduleName, CommonTokenStream tokenStream, CodeKind codeKind = CodeKind.SnippetCode, ParserMode parserMode = ParserMode.FallBackSllToLl);
    }
}
