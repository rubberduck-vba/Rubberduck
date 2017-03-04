using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class IndexExpression : BoundExpression
    {
        private readonly IBoundExpression _lExpression;
        private readonly ArgumentList _argumentList;

        public IndexExpression(
            Declaration referencedDeclaration, 
            ExpressionClassification classification, 
            ParserRuleContext context,
            IBoundExpression lExpression,
            ArgumentList argumentList)
            : base(referencedDeclaration, classification, context)
        {
            _lExpression = lExpression;
            _argumentList = argumentList;
        }

        public IBoundExpression LExpression
        {
            get
            {
                return _lExpression;
            }
        }

        public ArgumentList ArgumentList
        {
            get
            {
                return _argumentList;
            }
        }
    }
}
