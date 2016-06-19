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

        public Pattern(string expression, bool ignoreCase, bool global)
        {
            Flags = ignoreCase ? MatcherFlags.IgnoreCase : 0;
            Flags = global ? Flags | MatcherFlags.Global : Flags;

            RootExpression = RegularExpression.Parse(expression);
        }

    }

    [Flags]
    enum MatcherFlags
    {
        IgnoreCase = 1 << 0,
        Global     = 1 << 1,
    }
}
