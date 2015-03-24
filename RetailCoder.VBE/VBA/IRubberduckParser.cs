using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;

namespace Rubberduck.VBA
{
    public interface IRubberduckParser
    {
        /// <summary>
        /// Parses all code modules in specified project.
        /// </summary>
        /// <returns>Returns an <c>IParseTree</c> for each code module in the project; the qualified module name being the key.</returns>
        VBProjectParseResult Parse(VBProject vbProject);
    }
}