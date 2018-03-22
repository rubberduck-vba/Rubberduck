using Antlr4.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionValueResults
    {
        IUnreachableCaseInspectionValue GetValue(ParserRuleContext context);
        string GetTypeName(ParserRuleContext context);
        string GetValueString(ParserRuleContext context);
        bool Contains(ParserRuleContext context);
        void AddResult(ParserRuleContext context, IUnreachableCaseInspectionValue value);
        bool TryGetValue(ParserRuleContext context, out IUnreachableCaseInspectionValue value);
    }

    public class UnreachableCaseInspectionValueResults : IUnreachableCaseInspectionValueResults
    {
        private Dictionary<ParserRuleContext, IUnreachableCaseInspectionValue> _results;
        public UnreachableCaseInspectionValueResults()
        {
            _results = new Dictionary<ParserRuleContext, IUnreachableCaseInspectionValue>();
        }

        public IUnreachableCaseInspectionValue GetValue(ParserRuleContext context)
        {
            return _results[context];
        }

        public string GetTypeName(ParserRuleContext context)
        {
            return GetValue(context).TypeName;
        }

        public string GetValueString(ParserRuleContext context)
        {
            return GetValue(context).ValueText;
        }

        public bool Contains(ParserRuleContext context)
        {
            return _results.ContainsKey(context);
        }

        public void AddResult(ParserRuleContext context, IUnreachableCaseInspectionValue value)
        {
            _results.Add(context, value);
        }

        public bool TryGetValue(ParserRuleContext context, out IUnreachableCaseInspectionValue value)
        {
            return _results.TryGetValue(context, out value);
        }
    }

}
