
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionFactoryProvider
    {
        IParseTreeValueVisitorFactory CreateIParseTreeValueVisitorFactory();
        IParseTreeValueFactory CreateIParseTreeValueFactory();
        IRangeClauseFilterFactory CreateIRangeClauseFilterFactory();
        ISelectCaseStmtContextWrapperFactory CreateISelectStmtContextWrapperFactory();
        IRangeClauseContextWrapperFactory CreateIRangeClauseContextWrapperFactory();
    }

    public class UnreachableCaseInspectionFactoryProvider : IUnreachableCaseInspectionFactoryProvider
    {
        public IParseTreeValueVisitorFactory CreateIParseTreeValueVisitorFactory()
        {
            return new ParseTreeValueVisitorFactory();
        }

        public IParseTreeValueFactory CreateIParseTreeValueFactory()
        {
            return new ParseTreeValueFactory();
        }

        public IRangeClauseFilterFactory CreateIRangeClauseFilterFactory()
        {
            return new RangeClauseFilterFactory();
        }

        public ISelectCaseStmtContextWrapperFactory CreateISelectStmtContextWrapperFactory()
        {
            return new SelectCaseStmtContextWrapperFactory()
            {
                FactoryProvider = this
            };
        }

        public IRangeClauseContextWrapperFactory CreateIRangeClauseContextWrapperFactory()
        {
            return new RangeClauseContextWrapperFactory()
            {
                FactoryProvider = this
            };
        }
    }
}
