using Rubberduck.RegexAssistant.i18n;

namespace Rubberduck.RegexAssistant
{
    static class QuantifierExtensions
    {
        public static string HumanReadable(this Quantifier quant)
        {
            switch (quant.Kind)
            {
                case QuantifierKind.None:
                    return AssistantResources.Quantifier_None;
                case QuantifierKind.Wildcard:
                    if (quant.MaximumMatches == 1)
                    {
                        return AssistantResources.Quantifier_Optional;
                    }
                    if (quant.MinimumMatches == 0)
                    {
                        return AssistantResources.Quantifier_Asterisk;
                    }
                    return AssistantResources.Quantifer_Plus;
                case QuantifierKind.Expression:
                    if (quant.MaximumMatches == quant.MinimumMatches)
                    {
                        return string.Format(AssistantResources.Quantifier_Exact, quant.MinimumMatches);
                    }
                    if (quant.MaximumMatches == int.MaxValue)
                    {
                        return string.Format(AssistantResources.Quantifier_OpenRange, quant.MinimumMatches);
                    }
                    return string.Format(AssistantResources.Quantifier_ClosedRange, quant.MinimumMatches, quant.MaximumMatches);
            }
            return "";
        }
    }
}

    