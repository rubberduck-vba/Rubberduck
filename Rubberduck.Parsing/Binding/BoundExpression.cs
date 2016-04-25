using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public abstract class BoundExpression : IBoundExpression
    {
        private readonly Declaration _referencedDeclaration;
        private readonly ExpressionClassification _classification;
        private readonly ParserRuleContext _context;

        public BoundExpression(Declaration referencedDeclaration, ExpressionClassification classification, ParserRuleContext context)
        {
            _referencedDeclaration = referencedDeclaration;
            _classification = classification;
            _context = context;
        }

        public ExpressionClassification Classification
        {
            get
            {
                return _classification;
            }
        }

        public Declaration ReferencedDeclaration
        {
            get
            {
                return _referencedDeclaration;
            }
        }

        public ParserRuleContext Context
        {
            get
            {
                return _context;
            }
        }
    }
}
