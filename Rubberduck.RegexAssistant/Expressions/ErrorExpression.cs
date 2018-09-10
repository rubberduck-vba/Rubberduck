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

        public string Description => string.Format(AssistantResources.ExpressionDescription_ErrorExpression, _errorToken);

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>();

        public override string ToString() => $"Error Expression for {_errorToken}";
        public override bool Equals(object obj)
        {
            var expression = obj as ErrorExpression;
            return expression != null &&
                   _errorToken == expression._errorToken;
        }
        public override int GetHashCode()
        {
            return 330132629 + EqualityComparer<string>.Default.GetHashCode(_errorToken);
        }
    }
}