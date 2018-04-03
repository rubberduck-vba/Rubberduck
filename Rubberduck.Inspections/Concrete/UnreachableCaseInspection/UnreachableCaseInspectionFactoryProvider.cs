
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionFactoryProvider
    {
        IParseTreeValueVisitorFactory CreateIUCIParseTreeValueVisitorFactory();
        IParseTreeValueFactory CreateIUCIValueFactory();
        IRangeClauseFilterFactory CreateIUCIRangeClauseFilterFactory();
        ISelectCaseStmtContextWrapperFactory CreateUnreachableCaseInspectionSelectStmtFactory();
        IRangeClauseContextWrapperFactory CreateUnreachableCaseInspectionRangeFactory();
    }

    public class UnreachableCaseInspectionFactoryProvider : IUnreachableCaseInspectionFactoryProvider
    {
        public IParseTreeValueVisitorFactory CreateIUCIParseTreeValueVisitorFactory()
        {
            return new ParseTreeValueVisitorFactory();
        }

        public IParseTreeValueFactory CreateIUCIValueFactory()
        {
            return new ParseTreeValueFactory();
        }

        public IRangeClauseFilterFactory CreateIUCIRangeClauseFilterFactory()
        {
            return new RangeClauseFilterFactory();
        }

        public ISelectCaseStmtContextWrapperFactory CreateUnreachableCaseInspectionSelectStmtFactory()
        {
            return new SelectCaseStmtContextWrapperFactory()
            {
                FactoryProvider = this
            };
        }

        public IRangeClauseContextWrapperFactory CreateUnreachableCaseInspectionRangeFactory()
        {
            return new RangeClauseContextWrapperFactory()
            {
                FactoryProvider = this
            };
        }
    }
}
