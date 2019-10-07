using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class OutputListDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly List<IExpressionBinding> _itemBindings;

        public OutputListDefaultBinding(
            ParserRuleContext context,
            List<IExpressionBinding> itemBindings)
        {
            _context = context;
            _itemBindings = itemBindings;
        }

        public IBoundExpression Resolve()
        {
            var itemExpressions = new List<IBoundExpression>();
            foreach (var itemBinding in _itemBindings)
            {
                itemExpressions.Add(itemBinding.Resolve());
            };
            return new OutputListExpression(_context, itemExpressions);
        }
    }
}