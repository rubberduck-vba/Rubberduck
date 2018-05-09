using Rubberduck.UnitTesting;

namespace Rubberduck.ComClientLibrary.UnitTesting
{
    public class FakesProviderFactory : IFakesFactory 
    {
        public IFakes Create()
        {
            return new FakesProvider();
        }
    }
}
