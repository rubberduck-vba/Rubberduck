using System;
using System.Linq;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class LivelinessExpression : Expression
    {
        private readonly IExpression _isAlive;
        private readonly IExpression _code;

        public LivelinessExpression(IExpression isAlive, IExpression code)
        {
            _isAlive = isAlive;
            _code = code;
        }

        public override IValue Evaluate()
        {
            var isAlive = _isAlive.Evaluate().AsBool;
            var code = _code.Evaluate().AsString;
            return isAlive ? new StringValue(code) : new StringValue(MarkAsDead(code));
        }

        private string MarkAsDead(string code)
        {
            var hasNewLine = code.EndsWith(Environment.NewLine);
            // Remove parsed new line.
            code = code.Substring(0, code.Length - Environment.NewLine.Length);
            var lines = code.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var result = string.Join(Environment.NewLine, lines.Select(_ => string.Empty));
            if (hasNewLine)
            {
                result += Environment.NewLine;
            }
            return result;
        }
    }
}
