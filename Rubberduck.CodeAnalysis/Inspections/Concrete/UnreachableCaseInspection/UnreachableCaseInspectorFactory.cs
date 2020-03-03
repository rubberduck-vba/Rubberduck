using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectorFactory
    {
        IUnreachableCaseInspector Create(
            QualifiedModuleName module,
            VBAParser.SelectCaseStmtContext selectStmt, 
            IParseTreeVisitorResults results, 
            Func<string,QualifiedModuleName,ParserRuleContext,string> func = null);
    }

    public class UnreachableCaseInspectorFactory : IUnreachableCaseInspectorFactory
    {
        private readonly IParseTreeValueFactory _valueFactory;

        public UnreachableCaseInspectorFactory(IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IUnreachableCaseInspector Create(QualifiedModuleName module, VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results, Func<string, QualifiedModuleName, ParserRuleContext, string> func = null)
        {
            return new UnreachableCaseInspector(module, selectStmt, results, _valueFactory, func);
        }
    }
}
