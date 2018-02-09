using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class ContextValueResults<T> where T : IComparable<T>
    {
        private ContextExtents<T> _extents;
        public ContextValueResults(Dictionary<ParserRuleContext, T> values, Dictionary<ParserRuleContext, ParseTreeValue> unresolved)
        {
            _valueResolvedContexts = values;
            _unResolvedContexts = unresolved;
            _extents = new ContextExtents<T>();
        }

        private Dictionary<ParserRuleContext, T> _valueResolvedContexts;
        private Dictionary<ParserRuleContext, ParseTreeValue> _unResolvedContexts;
        public ContextExtents<T> Extents => _extents;
        public void SetExtents(T min, T max)
        {
            Extents.MinMax(min, max);
        }

        public Dictionary<ParserRuleContext, T> ValueResolvedContexts => _valueResolvedContexts;
        public Dictionary<ParserRuleContext, ParseTreeValue> UnresolvedContexts => _unResolvedContexts;
        public string EvaluationTypeName { set; get; } = string.Empty;
    }
}
