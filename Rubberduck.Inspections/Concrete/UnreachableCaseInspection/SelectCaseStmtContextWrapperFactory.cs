using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface ISelectCaseStmtContextWrapperFactory
    {
        ISelectCaseStmtContextWrapper Create(VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results);
        IUnreachableCaseInspectionFactoryProvider FactoryProvider { set; get; }
    }

    public class SelectCaseStmtContextWrapperFactory : ISelectCaseStmtContextWrapperFactory
    {
        public ISelectCaseStmtContextWrapper Create(VBAParser.SelectCaseStmtContext selectStmt, IParseTreeVisitorResults results)
        {
            return new SelectCaseStmtContextWrapper(selectStmt, results, FactoryProvider);
        }

        public IUnreachableCaseInspectionFactoryProvider FactoryProvider { set; get; }
    }

}
