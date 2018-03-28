using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionRangeFactory
    {
        IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, IUCIValueResults results);
        IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, string typeName, IUCIValueResults results);
        IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }

    public class UnreachableCaseInspectionRangeFactory : IUnreachableCaseInspectionRangeFactory
    {
        public IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, IUCIValueResults results)
        {
            return new UnreachableCaseInspectionRange(range, results, FactoryFactory);
        }

        public IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, string typeName, IUCIValueResults results)
        {
            return new UnreachableCaseInspectionRange(range, typeName, results, FactoryFactory);
        }

        public IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }
}
