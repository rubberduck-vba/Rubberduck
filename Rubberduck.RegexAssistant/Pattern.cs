using Rubberduck.RegexAssistant.i18n;
using System;

namespace Rubberduck.RegexAssistant
{
    public class Pattern : IDescribable
    {
        public IRegularExpression RootExpression;
        readonly MatcherFlags Flags;

        private readonly string _description;
        public string Description(bool spellOutWhitespace) => _description;

        public Pattern(string expression, bool ignoreCase = false, bool global = false, bool spellOutWhitespace = false)
        {
            if (expression == null)
            {
                throw new ArgumentNullException();
            }

            Flags = ignoreCase ? MatcherFlags.IgnoreCase : 0;
            Flags = global ? Flags | MatcherFlags.Global : Flags;

            AnchoredAtEnd = expression[expression.Length - 1].Equals('$');
            AnchoredAtStart = expression[0].Equals('^');

            var start = AnchoredAtStart ? 1 : 0;
            var end = (AnchoredAtEnd ? 1 : 0) + start;

            _spellOutWhiteSpace = spellOutWhitespace;
            RootExpression = VBRegexParser.Parse(expression.Substring(start, expression.Length - end), _spellOutWhiteSpace);
            _description = AssembleDescription();
        }

        private readonly bool _spellOutWhiteSpace;

        private string AssembleDescription()
        {
            var result = string.Empty;
            result += CasingDescription;
            result += StartAnchorDescription;
            result += RootExpression.Description(_spellOutWhiteSpace);
            result += EndAnchorDescription;
            return result;
        }

        public string StartAnchorDescription
        {
            get
            {
                if (AnchoredAtStart)
                {
                    return Flags.HasFlag(MatcherFlags.Global) 
                        ? AssistantResources.PatternDescription_AnchorStart_GlobalEnabled 
                        : AssistantResources.PatternDescription_AnchorStart;
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
                    return Flags.HasFlag(MatcherFlags.Global) 
                        ? AssistantResources.PatternDescription_AnchorEnd_GlobalEnabled 
                        : AssistantResources.PatternDescription_AnchorEnd;
                }
                return string.Empty;
            }
        }

        public bool IgnoreCase => Flags.HasFlag(MatcherFlags.IgnoreCase);
        public bool Global => Flags.HasFlag(MatcherFlags.Global);
        public bool AnchoredAtStart { get; }

        public bool AnchoredAtEnd { get; }

        public string CasingDescription => IgnoreCase ? AssistantResources.PatternDescription_IgnoreCase : string.Empty;
    }

    [Flags]
    internal enum MatcherFlags
    {
        IgnoreCase = 1 << 0,
        Global     = 1 << 1,
    }
}
