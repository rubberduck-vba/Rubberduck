
using Rubberduck.Resources.RegexAssistant;

namespace Rubberduck.RegexAssistant
{
    static class QuantifierExtensions
    {
        public static string HumanReadable(this Quantifier quant)
        {
            switch (quant.Kind)
            {
                case QuantifierKind.None:
                    return RegexAssistantUI.Quantifier_None;
                case QuantifierKind.Wildcard:
                    if (quant.MaximumMatches == 1)
                    {
                        return RegexAssistantUI.Quantifier_Optional;
                    }
                    if (quant.MinimumMatches == 0)
                    {
                        return RegexAssistantUI.Quantifier_Asterisk;
                    }
                    return RegexAssistantUI.Quantifer_Plus;
                case QuantifierKind.Expression:
                    if (quant.MaximumMatches == quant.MinimumMatches)
                    {
                        return string.Format(RegexAssistantUI.Quantifier_Exact, quant.MinimumMatches);
                    }
                    if (quant.MaximumMatches == int.MaxValue)
                    {
                        return string.Format(RegexAssistantUI.Quantifier_OpenRange, quant.MinimumMatches);
                    }
                    return string.Format(RegexAssistantUI.Quantifier_ClosedRange, quant.MinimumMatches, quant.MaximumMatches);
            }
            return "";
        }
    }
}

    