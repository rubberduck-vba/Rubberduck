using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;
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
        IEnumerable<VbModuleParseResult> Parse(VBProject vbProject);

        IEnumerable<CommentNode> ParseComments(VBComponent vbComponent);
    }

    public class VBParser : IRubberduckParser
    {
        public INode Parse(string projectName, string componentName, string code)
        {
            var result = Parse(code);
            var walker = new ParseTreeWalker();
            
            var listener = new VBTreeListener(projectName, componentName);
            walker.Walk(listener, result);

            return listener.Root;
        }

        public IParseTree Parse(string code)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VisualBasic6Lexer(input);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VisualBasic6Parser(tokens);
            
            var result = parser.startRule();
            return result;
        }

        public IEnumerable<VbModuleParseResult> Parse(VBProject project)
        {
            return project.VBComponents.Cast<VBComponent>()
                          .Select(component => new VbModuleParseResult(new QualifiedModuleName(project.Name, component.Name), 
                                               Parse(component.CodeModule.ToString()), ParseComments(component)));
        }

        public IEnumerable<CommentNode> ParseComments(VBComponent component)
        {
            return new List<CommentNode>();
            //todo: implement
        }
    }
}
