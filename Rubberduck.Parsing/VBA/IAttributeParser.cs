using System;
using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        (IParseTree tree, ITokenStream tokenStream) Parse(QualifiedModuleName module, CancellationToken cancellationToken);
    }
}
