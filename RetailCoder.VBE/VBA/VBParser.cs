using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
{
    public class VBParser
    {
        public SyntaxTreeNode Parse(string projectName, string componentName, string code)
        {
            var result = ParseInternal(code);

            return new ModuleNode(projectName, componentName, null, false);
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
