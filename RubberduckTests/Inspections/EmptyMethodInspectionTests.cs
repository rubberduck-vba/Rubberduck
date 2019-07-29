using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    class EmptyMethodInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void EmptyMethodBlock_InspectionName()
        {
            const string expectedName = nameof(EmptyMethodInspection);
            var inspection = new EmptyMethodInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
        }
    }
}
