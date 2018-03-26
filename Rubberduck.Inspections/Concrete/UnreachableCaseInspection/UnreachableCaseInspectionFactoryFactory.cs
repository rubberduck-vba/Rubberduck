
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionFactoryFactory
    {
        IUCIParseTreeValueVisitorFactory CreateIUCIParseTreeValueVisitorFactory();
        IUCIValueFactory CreateIUCIValueFactory();
        IUCIRangeClauseFilterFactory CreateIUCIRangeClauseFilterFactory();
        IUnreachableCaseInspectionSelectStmtFactory CreateUnreachableCaseInspectionSelectStmtFactory();
        IUnreachableCaseInspectionRangeFactory CreateUnreachableCaseInspectionRangeFactory();
    }

    public class UnreachableCaseInspectionFactoryFactory : IUnreachableCaseInspectionFactoryFactory
    {
        public IUCIParseTreeValueVisitorFactory CreateIUCIParseTreeValueVisitorFactory()
        {
            return new UCIParseTreeValueVisitorFactory();
        }

        public IUCIValueFactory CreateIUCIValueFactory()
        {
            return new UCIValueFactory();
        }

        public IUCIRangeClauseFilterFactory CreateIUCIRangeClauseFilterFactory()
        {
            return new UCIRangeClauseFilterFactory();
        }

        public IUnreachableCaseInspectionSelectStmtFactory CreateUnreachableCaseInspectionSelectStmtFactory()
        {
            return new UnreachableCaseInspectionSelectStmtFactory()
            {
                FactoryFactory = this
            };
        }

        public IUnreachableCaseInspectionRangeFactory CreateUnreachableCaseInspectionRangeFactory()
        {
            return new UnreachableCaseInspectionRangeFactory()
            {
                FactoryFactory = this
            };
        }
    }
}
