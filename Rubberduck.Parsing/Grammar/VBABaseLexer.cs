using Antlr4.Runtime;

namespace Rubberduck.Parsing.Grammar
{
    public abstract class VBABaseLexer : Lexer
    {
        public VBABaseLexer(ICharStream input) : base(input) { }

        #region Semantic predicate helper methods
        protected int CharAtRelativePosition(int i)
        {
            return _input.La(i);
        }

        protected bool IsChar(int actual, char expected)
        {
            return (char)actual == expected;
        }

        protected bool IsChar(int actual, params char[] expectedOptions)
        {
            char actualAsChar = (char)actual;
            foreach (char expected in expectedOptions)
            {
                if (actualAsChar == expected)
                {
                    return true;
                }
            }
            return false;
        }
        #endregion
    }
}
