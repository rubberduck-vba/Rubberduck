using System;
using Rubberduck.Resources.RegexAssistant;

namespace Rubberduck.RegexAssistant
{
    public class Pattern : IDescribable
    {
        public IRegularExpression RootExpression;
        private readonly MatcherFlags _flags;

        public string Description { get; }

        public Pattern(string expression, bool ignoreCase = false, bool global = false)
        {
            if (expression == null)
            {
                throw new ArgumentNullException();
            }

            _flags = ignoreCase ? MatcherFlags.IgnoreCase : 0;
            _flags = global ? _flags | MatcherFlags.Global : _flags;

            AnchoredAtEnd = expression[expression.Length - 1].Equals('$');
            AnchoredAtStart = expression[0].Equals('^');

            var start = AnchoredAtStart ? 1 : 0;
            var end = (AnchoredAtEnd ? 1 : 0) + start;
            RootExpression = RegularExpression.Parse(expression.Substring(start, expression.Length - end));
            Description = AssembleDescription();
        }

        private string AssembleDescription()
        {
            var result = string.Empty;
            result += CasingDescription;
            result += StartAnchorDescription;
            result += RootExpression.Description;
            result += EndAnchorDescription;
            return result;
        }

        public string StartAnchorDescription
        {
            get
            {
                if (AnchoredAtStart)
                {
                    return _flags.HasFlag(MatcherFlags.Global) 
                        ? RegexAssistantUI.PatternDescription_AnchorStart_GlobalEnabled 
                        : RegexAssistantUI.PatternDescription_AnchorStart;
                }
                return string.Empty;
            }
        }

        public string EndAnchorDescription
        {
            get
            {
                if (AnchoredAtEnd)
                {
                    return _flags.HasFlag(MatcherFlags.Global) 
                        ? RegexAssistantUI.PatternDescription_AnchorEnd_GlobalEnabled 
                        : RegexAssistantUI.PatternDescription_AnchorEnd;
                }
                return string.Empty;
            }
        }

        public bool IgnoreCase => _flags.HasFlag(MatcherFlags.IgnoreCase);
        public bool Global => _flags.HasFlag(MatcherFlags.Global);
        public bool AnchoredAtStart { get; }

        public bool AnchoredAtEnd { get; }

        public string CasingDescription => IgnoreCase ? RegexAssistantUI.PatternDescription_IgnoreCase : string.Empty;
    }

    [Flags]
    internal enum MatcherFlags
    {
        IgnoreCase = 1 << 0,
        Global     = 1 << 1,
    }
}
