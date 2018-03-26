using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIValueResults
    {
        IUCIValue GetValue(ParserRuleContext context);
        string GetTypeName(ParserRuleContext context);
        string GetValueText(ParserRuleContext context);
        bool Contains(ParserRuleContext context);
        bool TryGetValue(ParserRuleContext context, out IUCIValue value);
        void OnNewValueResult(object sender, ValueResultEventArgs e);
        bool Any();
    }

    public class UCIValueResults : IUCIValueResults
    {
        private Dictionary<ParserRuleContext, IUCIValue> _parseTreeValues;

        public UCIValueResults()
        {
            _parseTreeValues = new Dictionary<ParserRuleContext, IUCIValue>();
        }

        public IUCIValue GetValue(ParserRuleContext context)
        {
            if (context is null)
            {
                throw new ArgumentNullException();
            }
            return _parseTreeValues[context];
        }

        public string GetTypeName(ParserRuleContext context)
        {
            return GetValue(context).TypeName;
        }

        public string GetValueText(ParserRuleContext context)
        {
            return GetValue(context).ValueText;
        }

        public bool Contains(ParserRuleContext context)
        {
            return _parseTreeValues.ContainsKey(context);
        }

        public bool Any()
        {
            return _parseTreeValues.Count == 0;
        }

        public bool TryGetValue(ParserRuleContext context, out IUCIValue value)
        {
            return _parseTreeValues.TryGetValue(context, out value);
        }

        public void OnNewValueResult(object sender, ValueResultEventArgs e)
        {
            if (_parseTreeValues.ContainsKey(e.Context))
            {
                return;
            }
            _parseTreeValues.Add(e.Context, e.Value);
        }
    }
}
