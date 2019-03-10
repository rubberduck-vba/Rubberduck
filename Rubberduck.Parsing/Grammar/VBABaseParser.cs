using Antlr4.Runtime;
using System;
using System.Text.RegularExpressions;

namespace Rubberduck.Parsing.Grammar
{
    public abstract class VBABaseParser : Parser
    {
        public VBABaseParser(ITokenStream input) : base(input) { }

        #region Semantic predicate helper methods
        protected int TokenTypeAtRelativePosition(int i)
        {
            return _input.La(i);
        }

        protected IToken TokenAtRelativePosition(int i)
        {
            return _input.Lt(i);
        }

        protected string TextOf(IToken token)
        {
            return token.Text;
        }

        protected bool MatchesRegex(string text, string pattern)
        {
            return Regex.Match(text,pattern).Success;
        }

        protected bool EqualsStringIgnoringCase(string actual, string expected)
        {
            return actual.Equals(expected,StringComparison.OrdinalIgnoreCase);
        }

        protected bool EqualsStringIgnoringCase(string actual, params string[] expectedOptions)
        {
            foreach (string expected in expectedOptions)
            {
                if (actual.Equals(expected,StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        protected bool EqualsString(string actual, string expected)
        {
            return actual.Equals(expected,StringComparison.Ordinal);
        }

        protected bool EqualsString(string actual, params string[] expectedOptions)
        {
            foreach (string expected in expectedOptions)
            {
                if (actual.Equals(expected,StringComparison.Ordinal))
                {
                    return true;
                }
            }
            return false;
        }

        protected bool IsTokenType(int actual, params int[] expectedOptions)
        {
            foreach (int expected in expectedOptions)
            {
                if (actual == expected)
                {
                    return true;
                }
            }
            return false;
        }
        #endregion
    }
}
