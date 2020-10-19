using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal interface IParseTreeVisitorResults
    {
        IParseTreeValue GetValue(ParserRuleContext context);
        List<ParserRuleContext> GetChildResults(ParserRuleContext parent);
        string GetValueType(ParserRuleContext context);
        string GetToken(ParserRuleContext context);
        bool Contains(ParserRuleContext context);
        bool TryGetValue(ParserRuleContext context, out IParseTreeValue value);
        bool TryGetEnumMembers(VBAParser.EnumerationStmtContext enumerationStmtContext, out IReadOnlyList<EnumMember> enumMembers);
    }

    internal interface IMutableParseTreeVisitorResults : IParseTreeVisitorResults
    {
        void AddIfNotPresent(ParserRuleContext context, IParseTreeValue value);
        void AddEnumMember(VBAParser.EnumerationStmtContext enumerationStmtContext, EnumMember enumMember);
    }

    internal class ParseTreeVisitorResults : IMutableParseTreeVisitorResults
    {
        private readonly Dictionary<ParserRuleContext, IParseTreeValue> _parseTreeValues = new Dictionary<ParserRuleContext, IParseTreeValue>();
        private readonly Dictionary<VBAParser.EnumerationStmtContext, List<EnumMember>> _enumMembers = new Dictionary<VBAParser.EnumerationStmtContext, List<EnumMember>>();

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

        public bool TryGetEnumMembers(VBAParser.EnumerationStmtContext enumerationStmtContext, out IReadOnlyList<EnumMember> enumMembers)
        {
            if (!_enumMembers.TryGetValue(enumerationStmtContext, out var enumMemberList))
            {
                enumMembers = null;
                return false;
            }

            enumMembers = enumMemberList;
            return true;
        }

        public void AddEnumMember(VBAParser.EnumerationStmtContext enumerationStmtContext, EnumMember enumMember)
        {
            if (_enumMembers.TryGetValue(enumerationStmtContext, out var enumMemberList))
            {
                enumMemberList.Add(enumMember);
            }
            else
            {
                enumMemberList = new List<EnumMember>{enumMember};
                _enumMembers.Add(enumerationStmtContext, enumMemberList);
            }
        }
    }
}
