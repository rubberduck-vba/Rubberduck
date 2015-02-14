using System.Collections.Generic;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public interface IRubberduckParser
    {
        /// <summary>
        /// Parses specified code into a COM-visible code tree.
        /// </summary>
        /// <param name="projectName">The name of the VBA project the code belongs to.</param>
        /// <param name="componentName">The name of the VBA component (module) the code belongs to.</param>
        /// <param name="code">The code fragment or module content to be parsed.</param>
        /// <returns>Returns a COM-visible code tree.</returns>
        INode Parse(string projectName, string componentName, string code);

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