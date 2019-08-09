using Rubberduck.RegexAssistant.i18n;
using Rubberduck.VBEditor;
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

        public string Description(bool spellOutWhitespace) => string.Format(AssistantResources.ExpressionDescription_AlternativesExpression, Subexpressions.Count);

        public IList<IRegularExpression> Subexpressions { get; }

        public override string ToString() => $"Alternatives:{Subexpressions}";
        public override bool Equals(object obj) => obj is AlternativesExpression other && Subexpressions.Equals(other.Subexpressions);
        public override int GetHashCode() => HashCode.Compute(Subexpressions);
    }
}