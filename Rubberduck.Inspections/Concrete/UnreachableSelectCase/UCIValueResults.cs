using Antlr4.Runtime;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUCIValueResults
    {
        IUCIValue GetValue(ParserRuleContext context);
        string GetTypeName(ParserRuleContext context);
        string GetValueString(ParserRuleContext context);
        bool Contains(ParserRuleContext context);
        void AddResult(ParserRuleContext context, IUCIValue value);
        bool TryGetValue(ParserRuleContext context, out IUCIValue value);
    }

    public class UCIValueResults : IUCIValueResults
    {
        private Dictionary<ParserRuleContext, IUCIValue> _results;
        public UCIValueResults()
        {
            _results = new Dictionary<ParserRuleContext, IUCIValue>();
        }

        public IUCIValue GetValue(ParserRuleContext context)
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

        public void AddResult(ParserRuleContext context, IUCIValue value)
        {
            _results.Add(context, value);
        }

        public bool TryGetValue(ParserRuleContext context, out IUCIValue value)
        {
            return _results.TryGetValue(context, out value);
        }
    }

}
