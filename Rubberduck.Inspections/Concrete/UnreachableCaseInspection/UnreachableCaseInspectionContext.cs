using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public abstract class UnreachableCaseInspectionContext
    {
        protected readonly ParserRuleContext _context;
        protected IUCIValueResults _inspValues;
        private readonly IUCIRangeClauseFilterFactory _rangeFilterFactory;
        private readonly IUCIValueFactory _valueFactory;

        public UnreachableCaseInspectionContext(ParserRuleContext context, IUCIValueResults inspValues, IUnreachableCaseInspectionFactoryFactory factoryFactory) //, IUCIRangeClauseFilterFactory factory, IUCIValueFactory valueFactory)
        {
            _context = context;
            _rangeFilterFactory = factoryFactory.CreateIUCIRangeClauseFilterFactory();
            _valueFactory = factoryFactory.CreateIUCIValueFactory();
            _inspValues = inspValues;
        }

        protected IUCIValueFactory ValueFactory => _valueFactory;

        protected IUCIRangeClauseFilterFactory FilterFactory => _rangeFilterFactory;

        protected ParserRuleContext Context => _context;

        protected IUCIValueResults ParseTreeValueResults => _inspValues;
    }
}
