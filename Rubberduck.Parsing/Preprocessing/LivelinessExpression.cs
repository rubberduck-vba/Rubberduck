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
            bool isAlive = _isAlive.Evaluate().AsBool;
            var code = _code.Evaluate().AsString;
            if (isAlive)
            {
                return new StringValue(code);
            }
            else
            {
                return new StringValue(MarkAsDead(code));
            }
        }

        private string MarkAsDead(string code)
        {
            bool hasNewLine = false;
            if (code.EndsWith(Environment.NewLine))
            {
                hasNewLine = true;
            }
            // Remove parsed new line.
            code = code.TrimEnd('\r', '\n');
            var lines = code.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var result = string.Join(Environment.NewLine, lines.Select(_ => string.Empty));
            if (hasNewLine)
            {
                result += Environment.NewLine;
            }
            return result;
        }

        private string MarkLineAsDead(string line)
        {
            var result = string.Empty;
            if (line.EndsWith(Environment.NewLine))
            {
                result += Environment.NewLine;
            }
            return result;
        }
    }
}
