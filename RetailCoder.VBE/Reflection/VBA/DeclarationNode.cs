using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal class DeclarationNode : SyntaxTreeNode
    {
        public DeclarationNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match)
        {

        }

        IEnumerable<DeclarationNodeBase> _declarations;
        IEnumerable<DeclarationNodeBase> Declarations { get { return _declarations; } }
    }
}
