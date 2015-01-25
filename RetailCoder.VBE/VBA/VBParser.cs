using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public interface IRubberduckParser
    {
        /// <summary>
        /// Parses specified code into a code tree.
        /// </summary>
        /// <param name="projectName">The name of the VBA project the code belongs to.</param>
        /// <param name="componentName">The name of the VBA component (module) the code belongs to.</param>
        /// <param name="code">The code fragment or to be parsed.</param>
        /// <returns></returns>
        INode Parse(string projectName, string componentName, string code);
    }

    public class VBParser : IRubberduckParser
    {
        public INode Parse(string projectName, string componentName, string code)
        {
            var result = ParseInternal(code);
            var walker = new ParseTreeWalker();
            
            var listener = new VBTreeListener(projectName, componentName);
            walker.Walk(listener, result);

            return listener.Root;
        }

        private IParseTree ParseInternal(string code)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VisualBasic6Lexer(input);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VisualBasic6Parser(tokens);
            
            return parser.startRule();
        }
    }
}
