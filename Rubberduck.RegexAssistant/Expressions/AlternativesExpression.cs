using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;

namespace Rubberduck.RegexAssistant.Expressions
{
    public class AlternativesExpression : IRegularExpression
    {
        public AlternativesExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions ?? throw new ArgumentNullException();
        }

        public string Description => string.Format(AssistantResources.ExpressionDescription_AlternativesExpression, Subexpressions.Count);

        public IList<IRegularExpression> Subexpressions { get; }
    }
}