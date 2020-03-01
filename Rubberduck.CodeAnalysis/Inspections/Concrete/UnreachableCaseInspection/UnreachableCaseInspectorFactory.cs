using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectorFactory
    {
        IUnreachableCaseInspector Create(
            VBAParser.SelectCaseStmtContext selectStmt, 
            IParseTreeVisitorResults results, 
            Func<string,ParserRuleContext,string> func = null);
    }

    public class UnreachableCaseInspectorFactory : IUnreachableCaseInspectorFactory
    {
        private readonly IParseTreeValueFactory _valueFactory;

        public UnreachableCaseInspectorFactory(IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IUnreachableCaseInspector Create(VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results, Func<string, ParserRuleContext, string> func = null)
        {
            return new UnreachableCaseInspector(selectStmt, results, _valueFactory, func);
        }
    }
}
