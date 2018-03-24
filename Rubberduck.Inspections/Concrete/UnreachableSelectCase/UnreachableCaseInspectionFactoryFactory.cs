
namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionFactoryFactory
    {
        IUCIParseTreeValueVisitorFactory CreateVisitorFactory();
        IUCIValueFactory CreateValueFactory();
        IUCIRangeClauseFilterFactory CreateSummaryClauseFactory();
    }

    public class UnreachableCaseInspectionFactoryFactory : IUnreachableCaseInspectionFactoryFactory
    {
        public IUCIParseTreeValueVisitorFactory CreateVisitorFactory()
        {
            return new UCIParseTreeValueVisitorFactory();
        }

        public IUCIValueFactory CreateValueFactory()
        {
            return new UCIValueFactory();
        }

        public IUCIRangeClauseFilterFactory CreateSummaryClauseFactory()
        {
            return new UCIRangeClauseFilterFactory();
        }
    }

}
