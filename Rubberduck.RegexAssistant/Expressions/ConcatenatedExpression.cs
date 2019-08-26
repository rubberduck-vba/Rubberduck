using Rubberduck.RegexAssistant.i18n;
using Rubberduck.VBEditor;
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

        public string Description(bool spellOutWhitespace) => AssistantResources.ExpressionDescription_ConcatenatedExpression;

        public IList<IRegularExpression> Subexpressions { get; }

        public override string ToString() => $"Concatenated:{Subexpressions}";
        public override bool Equals(object obj) => obj is ConcatenatedExpression other && Subexpressions.Equals(other.Subexpressions);
        public override int GetHashCode() => HashCode.Compute(Subexpressions);
    }
}