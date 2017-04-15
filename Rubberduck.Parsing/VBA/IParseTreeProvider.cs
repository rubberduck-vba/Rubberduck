using Antlr4.Runtime.Tree;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    public interface IParseTreeProvider
    {
        List<KeyValuePair<QualifiedModuleName, IParseTree>> ParseTrees { get; }

        IParseTree GetParseTree(QualifiedModuleName module);
    }
}
