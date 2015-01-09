using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class ForLoopNode : CodeBlockNode
    {
        public ForLoopNode(Instruction instruction, string scope, Match match, IEnumerable<SyntaxTreeNode> nodes) 
            : base(instruction, scope, match, new[] {ReservedKeywords.Next}, null, nodes)
        {
        }

        /// <summary>
        /// Gets an <see cref="Expression"/> that returns the initial value of the loop counter.
        /// </summary>
        public Expression Lower { get { return new Expression(RegexMatch.Groups["lower"].Value); } }

        /// <summary>
        /// Gets an <see cref="Expression"/> that returns the final iteratable value of the loop counter.
        /// </summary>
        public Expression Upper { get { return new Expression(RegexMatch.Groups["upper"].Value); } }

        /// <summary>
        /// Gets an optional <see cref="Expression"/> that returns the <c>Step</c> argument of the <c>For</c> loop.
        /// Returns <c>null</c> if the argument is not specified.
        /// </summary>
        public Expression Step { get { return new Expression(RegexMatch.Groups["step"].Value); } }
    }
}