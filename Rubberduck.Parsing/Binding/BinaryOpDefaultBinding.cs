using Antlr4.Runtime;
using System.Diagnostics;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BinaryOpDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _left;
        private readonly IExpressionBinding _right;

        public BinaryOpDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding left,
            IExpressionBinding right)
        {
            _context = context;
            _left = left;
            _right = right;
        }

        public IBoundExpression Resolve()
        {
            var leftExpr = _left.Resolve();
            if (leftExpr == null)
            {
                return null;
            }
            var rightExpr = _right.Resolve();
            if (rightExpr == null)
            {
                return null;
            }
            return new BinaryOpExpression(null, _context, leftExpr, rightExpr);
        }
    }
}
