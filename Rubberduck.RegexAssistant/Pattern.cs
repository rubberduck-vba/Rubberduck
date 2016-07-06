using Rubberduck.RegexAssistant.i18n;
using System;

namespace Rubberduck.RegexAssistant
{
    public class Pattern : IDescribable
    {
        public IRegularExpression RootExpression;
        MatcherFlags Flags;

        private readonly bool _hasStartAnchor;
        private readonly bool _hasEndAnchor;
        private readonly string _description;

        public string Description
        {
            get
            {
                return _description;
            }
        }

        public Pattern(string expression, bool ignoreCase, bool global)
        {
            Flags = ignoreCase ? MatcherFlags.IgnoreCase : 0;
            Flags = global ? Flags | MatcherFlags.Global : Flags;

            _hasEndAnchor = expression[expression.Length - 1].Equals('$');
            _hasStartAnchor = expression[0].Equals('^');

            int start = _hasStartAnchor ? 1 : 0;
            int end = (_hasEndAnchor ? 1 : 0) + start;
            RootExpression = RegularExpression.Parse(expression.Substring(start, expression.Length - end));
            _description = AssembleDescription();
        }

        private string AssembleDescription()
        {
            string result = string.Empty;
            if (_hasStartAnchor)
            {
                result += Flags.HasFlag(MatcherFlags.Global) ? AssistantResources.PatternDescription_AnchorStart_GlobalEnabled : AssistantResources.PatternDescription_AnchorStart;
            }
            result += RootExpression.Description;
            if (_hasEndAnchor)
            {
                result += Flags.HasFlag(MatcherFlags.Global) ? AssistantResources.PatternDescription_AnchorEnd_GlobalEnabled : AssistantResources.PatternDescription_AnchorEnd;
            }
            return result;
        }
    }

    [Flags]
    enum MatcherFlags
    {
        IgnoreCase = 1 << 0,
        Global     = 1 << 1,
    }
}
