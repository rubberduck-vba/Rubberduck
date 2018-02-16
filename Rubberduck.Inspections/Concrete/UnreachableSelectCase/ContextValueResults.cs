using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    //public class ContextValueResults<T> where T : IComparable<T>
    //{
    //    private ContextExtents<T> _extents;
    //    public ContextValueResults(Dictionary<ParserRuleContext, T> values, Dictionary<ParserRuleContext, ParseTreeValue> unresolved)
    //    {
    //        _valueResolvedContexts = values;
    //        _variableContexts = unresolved;
    //        _extents = new ContextExtents<T>();
    //    }

    //    private bool ContainsBooleans => typeof(T) == typeof(bool);
    //    private bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);
    //    private Dictionary<ParserRuleContext, T> _valueResolvedContexts;
    //    private Dictionary<ParserRuleContext, ParseTreeValue> _variableContexts;

    //    public Dictionary<ParserRuleContext, T> ValueResolvedContexts => _valueResolvedContexts;
    //    public Dictionary<ParserRuleContext, ParseTreeValue> VariableContexts => _variableContexts;
    //    public ContextExtents<T> Extents => _extents;
    //    public void SetExtents(T min, T max)
    //    {
    //        Extents.MinMax(min, max);
    //    }
    //    public string EvaluationTypeName { set; get; } = string.Empty;
    //    public T TrueValue { set; get; }
    //    public T FalseValue { set; get; }
    //}
}
