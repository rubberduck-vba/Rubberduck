using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Grammar
{
    /// <summary>
    /// Base class for a declaration node.
    /// </summary>
    [ComVisible(false)]
    public abstract class DeclarationNode : SyntaxTreeNode
    {
        protected DeclarationNode(Instruction instruction, string scope, Match match, IEnumerable<SyntaxTreeNode> childNodes)
            : base(instruction, scope, match, childNodes)
        {

        }

        public string Accessibility
        {
            get
            {
                var specified = RegexMatch.Groups["keywords"].Value;
                if (string.IsNullOrEmpty(specified))
                {
                    return ReservedKeywords.Private;
                }

                return specified;
            }
        }
    }
}
