
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIValueFactory
    {
        IUCIValue Create(string valueToken, string declaredTypeName = null);
    }

    public class UCIValueFactory : IUCIValueFactory
    {
        public IUCIValue Create(string valueToken, string declaredTypeName = null)
        {
            return new UCIValue(valueToken, declaredTypeName);
        }
    }
}
