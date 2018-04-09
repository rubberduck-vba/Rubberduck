using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IRangeClauseContextWrapperFactory
    {
        IRangeClauseContextWrapper Create(VBAParser.RangeClauseContext range, IParseTreeVisitorResults results);
        IRangeClauseContextWrapper Create(VBAParser.RangeClauseContext range, string typeName, IParseTreeVisitorResults results);
        IUnreachableCaseInspectionFactoryProvider FactoryProvider { set; get; }
    }

    public class RangeClauseContextWrapperFactory : IRangeClauseContextWrapperFactory
    {
        public IRangeClauseContextWrapper Create(VBAParser.RangeClauseContext range, IParseTreeVisitorResults results)
        {
            return new RangeClauseContextWrapper(range, results, FactoryProvider);
        }

        public IRangeClauseContextWrapper Create(VBAParser.RangeClauseContext range, string typeName, IParseTreeVisitorResults results)
        {
            return new RangeClauseContextWrapper(range, typeName, results, FactoryProvider);
        }

        public IUnreachableCaseInspectionFactoryProvider FactoryProvider { set; get; }
    }
}
