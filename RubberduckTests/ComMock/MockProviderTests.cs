using System.Runtime.InteropServices;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    public class MockProviderTests
    {
        [Test]
        [Category("ComMocks")]
        public void MockProvider_Returns_Correct_Mock()
        {
            var provider = new MockProvider();
            var mock = provider.Mock("Excel.Application");
            var obj = mock.Object;

            Assert.IsTrue(Marshal.IsComObject(obj));
        }
    }
}
