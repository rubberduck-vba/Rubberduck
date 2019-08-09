using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;

namespace Rubberduck.RegexAssistant.Expressions
{
    public class ErrorExpression : IRegularExpression
    {
        private readonly string _errorToken;

        public ErrorExpression(string errorToken)
        {
            _errorToken = errorToken ?? throw new ArgumentNullException();
        }

        public string Description(bool spellOutWhitespace) => string.Format(AssistantResources.ExpressionDescription_ErrorExpression, _errorToken);

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>();

        public override string ToString() => $"Error Expression for {_errorToken}";
        public override bool Equals(object obj) => obj is ErrorExpression other && _errorToken == other._errorToken;
        public override int GetHashCode() => _errorToken.GetHashCode();
    }
}