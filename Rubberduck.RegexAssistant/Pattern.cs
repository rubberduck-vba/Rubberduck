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

        public Pattern(string expression, bool ignoreCase = false, bool global = false)
        {
            if (expression == null)
            {
                throw new ArgumentNullException();
            }

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

        public bool IgnoreCase { get { return Flags.HasFlag(MatcherFlags.IgnoreCase); } }
        public bool Global { get { return Flags.HasFlag(MatcherFlags.Global); } }
        public bool AnchoredAtStart { get { return _hasStartAnchor; } }
        public bool AnchoredAtEnd { get { return _hasEndAnchor; } }

        public string CasingDescription { get
            {
                return IgnoreCase ? AssistantResources.PatternDescription_IgnoreCase : string.Empty;
            }
        }
    }

    [Flags]
    enum MatcherFlags
    {
        IgnoreCase = 1 << 0,
        Global     = 1 << 1,
    }
}
