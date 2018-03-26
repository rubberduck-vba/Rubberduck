using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionRangeFactory
    {
        IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, IUCIValueResults results); //, IUnreachableCaseInspectionFactoryFactory factoryFactory);
        IUnreachableCaseInspectionRange Create(string typeName, VBAParser.RangeClauseContext range, IUCIValueResults results);  //, IUnreachableCaseInspectionFactoryFactory factoryFactory);
        IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }

    public class UnreachableCaseInspectionRangeFactory : IUnreachableCaseInspectionRangeFactory
    {
        public IUnreachableCaseInspectionRange Create(VBAParser.RangeClauseContext range, IUCIValueResults results) //, IUnreachableCaseInspectionFactoryFactory factoryFactory)
        {
            return new UnreachableCaseInspectionRange(range, results, FactoryFactory);
        }

        public IUnreachableCaseInspectionRange Create(string typeName, VBAParser.RangeClauseContext range, IUCIValueResults results) //, IUnreachableCaseInspectionFactoryFactory factoryFactory)
        {
            return new UnreachableCaseInspectionRange(range, results, FactoryFactory)
            {
                EvaluationTypeName = typeName
            };
        }

        public IUnreachableCaseInspectionFactoryFactory FactoryFactory { set; get; }
    }
}
