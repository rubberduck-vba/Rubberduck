using System.Globalization;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class HexNumberLiteralExpression : Expression
    {
        private readonly IExpression _tokenText;

        public HexNumberLiteralExpression(IExpression tokenText)
        {
            _tokenText = tokenText;
        }

        public override IValue Evaluate()
        {
            string literal = _tokenText.Evaluate().AsString;
            literal = literal.Replace("&H", "")
                .Replace("&", "")
                .Replace("%", "")
                .Replace("^", "");
            var number = (decimal)int.Parse(literal, NumberStyles.HexNumber);
            return new DecimalValue(number);
        }
    }
}
