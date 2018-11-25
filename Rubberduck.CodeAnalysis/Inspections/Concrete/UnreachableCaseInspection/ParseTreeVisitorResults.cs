using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeVisitorResults
    {
        IParseTreeValue GetValue(ParserRuleContext context);
        List<ParserRuleContext> GetChildResults(ParserRuleContext parent);
        string GetValueType(ParserRuleContext context);
        string GetToken(ParserRuleContext context);
        bool Contains(ParserRuleContext context);
        bool TryGetValue(ParserRuleContext context, out IParseTreeValue value);
        void OnNewValueResult(object sender, ValueResultEventArgs e);
    }

    public class ParseTreeVisitorResults : IParseTreeVisitorResults
    {
        private Dictionary<ParserRuleContext, IParseTreeValue> _parseTreeValues;

        public ParseTreeVisitorResults()
        {
            _parseTreeValues = new Dictionary<ParserRuleContext, IParseTreeValue>();
        }

        public IParseTreeValue GetValue(ParserRuleContext context)
        {
            if (context is null)
            {
                throw new ArgumentNullException();
            }
            return _parseTreeValues[context];
        }

        public List<ParserRuleContext> GetChildResults(ParserRuleContext parent)
        {
            var results = new List<ParserRuleContext>();
            if (!(parent is null) || Contains(parent))
            {
                foreach (var child in parent.children)
                {
                    if (child is ParserRuleContext prCtxt  && Contains(prCtxt))
                    {
                        results.Add(prCtxt);
                    }
                }
            }
            return results;
        }

        public string GetValueType(ParserRuleContext context)
        {
            return GetValue(context).ValueType;
        }

        public string GetToken(ParserRuleContext context)
        {
            return GetValue(context).Token;
        }

        public bool Contains(ParserRuleContext context)
        {
            return _parseTreeValues.ContainsKey(context);
        }

        public bool TryGetValue(ParserRuleContext context, out IParseTreeValue value)
        {
            return _parseTreeValues.TryGetValue(context, out value);
        }

        public void OnNewValueResult(object sender, ValueResultEventArgs e)
        {
            if (!_parseTreeValues.ContainsKey(e.Context))
            {
                _parseTreeValues.Add(e.Context, e.Value);
            }
        }
    }
}
