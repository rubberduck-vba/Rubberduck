using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionRangeFactory
    {
        IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, IUCIValueResults results);
        IUnreachableCaseInspectionRange Create(string typeName, VBAParser.RangeClauseContext range, IUCIValueResults results);
        IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }

    public class UnreachableCaseInspectionRangeFactory : IUnreachableCaseInspectionRangeFactory
    {
        public IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, IUCIValueResults results)
        {
            return new UnreachableCaseInspectionRange(range, results, FactoryFactory);
        }

        public IUnreachableCaseInspectionRange Create(string typeName, VBAParser.RangeClauseContext range, IUCIValueResults results)
        {
            return new UnreachableCaseInspectionRange(range, results, FactoryFactory)
            {
                EvaluationTypeName = typeName
            };
        }

        public IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }
}
