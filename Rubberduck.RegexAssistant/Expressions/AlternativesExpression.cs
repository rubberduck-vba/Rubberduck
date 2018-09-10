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

        public override string ToString() => $"Aternatives:{Subexpressions.ToString()}";
        public override bool Equals(object obj)
        {
            var expression = obj as AlternativesExpression;
            return expression != null &&
                   EqualityComparer<IList<IRegularExpression>>.Default.Equals(Subexpressions, expression.Subexpressions);
        }
        public override int GetHashCode()
        {
            return 1015294936 + EqualityComparer<IList<IRegularExpression>>.Default.GetHashCode(Subexpressions);
        }
    }
}