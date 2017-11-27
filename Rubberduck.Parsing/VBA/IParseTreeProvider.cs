using Antlr4.Runtime.Tree;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    public interface IParseTreeProvider
    {
        List<KeyValuePair<QualifiedModuleName, IParseTree>> ParseTrees { get; }
        List<KeyValuePair<QualifiedModuleName, IParseTree>> AttributeParseTrees { get; }
        IParseTree GetParseTree(QualifiedModuleName module, ParsePass pass);
    }
}
