using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionSelectStmtFactory
    {
        IUnreachableCaseInspectionSelectStmt Create(VBAParser.SelectCaseStmtContext selectStmt, IUCIValueResults results);
        IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }

    public class UnreachableCaseInspectionSelectStmtFactory : IUnreachableCaseInspectionSelectStmtFactory
    {
        public IUnreachableCaseInspectionSelectStmt Create(VBAParser.SelectCaseStmtContext selectStmt, IUCIValueResults results)
        {
            return new UnreachableCaseInspectionSelectStmt(selectStmt, results, FactoryFactory);
        }

        public IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }

}
