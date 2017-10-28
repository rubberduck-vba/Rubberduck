using System;
using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        (IParseTree tree, ITokenStream tokenStream, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes) Parse(IVBComponent component, CancellationToken cancellationToken);
    }
}
