using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    // todo: handle multiple declarations on single instruction.

    internal class ConstantNode : DeclarationNode
    {
        public ConstantNode(string scope, Match match, string comment)
            : base(scope, match, comment)
        { }

        /// <summary>
        /// Gets the constant's value. Strings include delimiting quotes.
        /// </summary>
        public string Value
        {
            get
            {
                return RegexMatch.Groups["value"].Value;
            }
        }
    }
}
