using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectorFactory
    {
        IUnreachableCaseInspector Create(VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results, IParseTreeValueFactory valueFactory);
    }

    public class UnreachableCaseInspectorFactory : IUnreachableCaseInspectorFactory
    {
        public IUnreachableCaseInspector Create(VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results, IParseTreeValueFactory valueFactory)
        {
            return new UnreachableCaseInspector(selectStmt, results, valueFactory);
        }
    }

}
