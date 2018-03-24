
namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
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
