using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

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
        IReadOnlyList<EnumMember> EnumMembers { get; }
    }

    public interface IMutableParseTreeVisitorResults : IParseTreeVisitorResults
    {
        void AddIfNotPresent(ParserRuleContext context, IParseTreeValue value);
        void Add(EnumMember enumMember);
    }

    public class ParseTreeVisitorResults : IMutableParseTreeVisitorResults
    {
        private readonly Dictionary<ParserRuleContext, IParseTreeValue> _parseTreeValues = new Dictionary<ParserRuleContext, IParseTreeValue>();
        private readonly List<EnumMember> _enumMembers = new List<EnumMember>();

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
            if (parent is null)
            {
                return new List<ParserRuleContext>();
            }

            return parent.children
                .OfType<ParserRuleContext>()
                .Where(Contains)
                .ToList();
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

        public void AddIfNotPresent(ParserRuleContext context, IParseTreeValue value)
        {
            if (!_parseTreeValues.ContainsKey(context))
            {
                _parseTreeValues.Add(context, value);
            }
        }

        public IReadOnlyList<EnumMember> EnumMembers => _enumMembers;
        public void Add(EnumMember enumMember)
        {
            _enumMembers.Add(enumMember);
        }
    }
}
