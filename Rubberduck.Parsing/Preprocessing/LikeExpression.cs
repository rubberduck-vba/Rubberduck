using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class LikeExpression : Expression
    {
        private readonly IExpression _expression;
        private readonly IExpression _pattern;

        public LikeExpression(IExpression expression, IExpression pattern)
        {
            _expression = expression;
            _pattern = pattern;
        }

        public override IValue Evaluate()
        {
            var expr = _expression.Evaluate();
            var pattern = _pattern.Evaluate();
            if (expr == null || pattern == null)
            {
                return null;
            }
            var exprStr = expr.AsString;
            var patternStr = pattern.AsString;
            var parser = new VBALikePatternParser();
            var likePattern = parser.Parse(patternStr);
            var regex = TranslateToNETRegex(likePattern);
            return new BoolValue(Regex.IsMatch(exprStr, regex));
        }

        private string TranslateToNETRegex(VBALikeParser.LikePatternStringContext likePattern)
        {
            StringBuilder regexStr = new StringBuilder();
            foreach (var element in likePattern.likePatternElement())
            {
                if (element.likePatternChar() != null)
                {
                    regexStr.Append(element.likePatternChar().GetText());
                }
                else if (element.QUESTIONMARK() != null)
                {
                    regexStr.Append(".");
                }
                else if (element.HASH() != null)
                {
                    regexStr.Append(@"\d");
                }
                else if (element.STAR() != null)
                {
                    regexStr.Append(@".*?");
                }
                else
                {
                    var charlist = element.likePatternCharlist().GetText();
                    if (charlist.StartsWith("[!"))
                    {
                        charlist = "[^" + charlist.Substring(2);
                    }
                    regexStr.Append(charlist);
                }
            }
            // Full string match, e.g. "abcd" should NOT match "a.c"
            var regex = "^" + regexStr.ToString() + "$";
            return regex;
        }
    }
}
