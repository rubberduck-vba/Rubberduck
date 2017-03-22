namespace Rubberduck.UnitTesting
{
    public interface IFakesProviderFactory
    {
        FakesProvider GetFakesProvider();
    }

    public class FakesProviderFactory : IFakesProviderFactory
    {
        public FakesProvider GetFakesProvider()
        {
            return new FakesProvider(); 
        }
    }
}
