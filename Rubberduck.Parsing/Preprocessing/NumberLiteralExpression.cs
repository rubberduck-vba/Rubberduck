using System.Globalization;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class NumberLiteralExpression : Expression
    {
        private readonly IExpression _tokenText;

        public NumberLiteralExpression(IExpression tokenText)
        {
            _tokenText = tokenText;
        }

        public override IValue Evaluate()
        {
            string literal = _tokenText.Evaluate().AsString;
            var number = decimal.Parse(literal
                .Replace("%", "")
                .Replace("&", "")
                .Replace("^", "")
                .Replace("!", "")
                .Replace("#", "")
                .Replace("@", "")
                , NumberStyles.Float, CultureInfo.InvariantCulture);
            return new DecimalValue(number);
        }
    }
}
