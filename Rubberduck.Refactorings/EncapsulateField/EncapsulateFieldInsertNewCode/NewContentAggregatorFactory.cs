
namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface INewContentAggregatorFactory
    {
        INewContentAggregator Create();
    }

    public class NewContentAggregatorFactory : INewContentAggregatorFactory
    {
        public INewContentAggregator Create()
        {
            return new NewContentAggregator();
        }
    }
}
