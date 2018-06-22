using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectorFactory
    {
        IUnreachableCaseInspector Create(VBAParser.SelectCaseStmtContext selectStmt, 
            IParseTreeVisitorResults results, 
            IParseTreeValueFactory valueFactory,
            Func<string,ParserRuleContext,string> func = null);
    }

    public class UnreachableCaseInspectorFactory : IUnreachableCaseInspectorFactory
    {
        public IUnreachableCaseInspector Create(VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results, IParseTreeValueFactory valueFactory, Func<string, ParserRuleContext, string> func = null)
        {
            return new UnreachableCaseInspector(selectStmt, results, valueFactory, func);
        }
    }
}
