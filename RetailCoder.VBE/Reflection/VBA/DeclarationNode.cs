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
        public DeclarationNode(string scope, Match match, string comment)
            : base(scope, match, comment, true)
        {

        }


        IEnumerable<DeclarationNodeBase> _declarations;
        IEnumerable<DeclarationNodeBase> Declarations { get { return _declarations; } }
    }
}
