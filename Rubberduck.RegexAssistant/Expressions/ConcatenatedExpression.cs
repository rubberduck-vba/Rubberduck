using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;

namespace Rubberduck.RegexAssistant.Expressions
{
    public class ConcatenatedExpression : IRegularExpression
    {
        public ConcatenatedExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions ?? throw new ArgumentNullException();
        }

        public string Description => AssistantResources.ExpressionDescription_ConcatenatedExpression;

        public IList<IRegularExpression> Subexpressions { get; }

        public override string ToString() => $"Concatenated:{Subexpressions.ToString()}";
        public override bool Equals(object obj)
        {
            var expression = obj as ConcatenatedExpression;
            return expression != null &&
                   EqualityComparer<IList<IRegularExpression>>.Default.Equals(Subexpressions, expression.Subexpressions);
        }
        public override int GetHashCode()
        {
            return 1015294936 + EqualityComparer<IList<IRegularExpression>>.Default.GetHashCode(Subexpressions);
        }
    }
}