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
        public ConstantNode(Match match)
            : base(match)
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

        /// <summary>
        /// Gets a value indicating whether the value is accessible beyond the scope where the constant is declared.
        /// </summary>
        public bool IsPublicScope
        {
            get
            {
                return new[] { ReservedKeywords.Public, ReservedKeywords.Global }.Contains(RegexMatch.Groups[0].Value);
            }
        }
    }
}
