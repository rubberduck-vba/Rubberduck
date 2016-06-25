using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.RegexAssistant
{
    class Pattern
    {
        IRegularExpression RootExpression;
        MatcherFlags Flags;

        private readonly bool _hasStartAnchor;
        private readonly bool _hasEndAnchor;

        public Pattern(string expression, bool ignoreCase, bool global)
        {
            Flags = ignoreCase ? MatcherFlags.IgnoreCase : 0;
            Flags = global ? Flags | MatcherFlags.Global : Flags;

            _hasEndAnchor = expression[expression.Length - 1].Equals('$');
            _hasStartAnchor = expression[0].Equals('^');

            int start = _hasStartAnchor ? 1 : 0;
            int end = (_hasEndAnchor ? 1 : 0) + start + 1;
            RootExpression = RegularExpression.Parse(expression.Substring(start, expression.Length - end));
        }

    }

    [Flags]
    enum MatcherFlags
    {
        IgnoreCase = 1 << 0,
        Global     = 1 << 1,
    }
}
