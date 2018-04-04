using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public abstract class ContextWrapperBase
    {
        protected readonly ParserRuleContext _context;
        protected IParseTreeVisitorResults _inspValues;
        private readonly IRangeClauseFilterFactory _rangeFilterFactory;
        private readonly IParseTreeValueFactory _valueFactory;

        public ContextWrapperBase(ParserRuleContext context, IParseTreeVisitorResults inspValues, IUnreachableCaseInspectionFactoryProvider factoryFactory)
        {
            _context = context;
            _rangeFilterFactory = factoryFactory.CreateIRangeClauseFilterFactory();
            _valueFactory = factoryFactory.CreateIParseTreeValueFactory();
            _inspValues = inspValues;
        }

        protected IParseTreeValueFactory ValueFactory => _valueFactory;

        protected IRangeClauseFilterFactory FilterFactory => _rangeFilterFactory;

        public ParserRuleContext Context => _context;

        protected IParseTreeVisitorResults ParseTreeValueResults => _inspValues;
    }
}
