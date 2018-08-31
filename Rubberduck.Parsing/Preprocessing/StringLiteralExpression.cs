namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class StringLiteralExpression : Expression
    {
        private readonly IExpression _tokenText;

        public StringLiteralExpression(IExpression tokenText)
        {
            _tokenText = tokenText;
        }

        public override IValue Evaluate()
        {
            var str = _tokenText.Evaluate().AsString;
            // Remove quotes
            str = str.Substring(1, str.Length - 2);
            return new StringValue(str);
        }
    }
}
