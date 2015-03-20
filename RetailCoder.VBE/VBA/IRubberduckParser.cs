using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public interface IRubberduckParser
    {
        /// <summary>
        /// Parses specified code into a parse tree.
        /// </summary>
        /// <param name="code">The code fragment or module content to be parsed.</param>
        /// <returns>Returns a parse tree representing the parsed code.</returns>
        IParseTree Parse(string code);

        /// <summary>
        /// Parses all code modules in specified project.
        /// </summary>
        /// <returns>Returns an <c>IParseTree</c> for each code module in the project; the qualified module name being the key.</returns>
        IEnumerable<VBComponentParseResult> Parse(VBProject vbProject);

        VBComponentParseResult Parse(VBComponent component);

        IEnumerable<CommentNode> ParseComments(VBComponent vbComponent);
    }
}