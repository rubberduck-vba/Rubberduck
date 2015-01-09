using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public class ForEachLoopNode : CodeBlockNode
    {
        public ForEachLoopNode(Instruction instruction, string scope, Match match, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, scope, match, new[] {ReservedKeywords.Next}, null, nodes)
        {
        }

        public Identifier Identifier { get { return new Identifier(base.Scope, RegexMatch.Groups["identifier"].Value, string.Empty);} }

        /// <summary>
        /// Gets an <see cref="Expression"/> that returns the object being iterated.
        /// </summary>
        public Expression Expression { get { return new Expression(RegexMatch.Groups["expression"].Value); } }
    }
}