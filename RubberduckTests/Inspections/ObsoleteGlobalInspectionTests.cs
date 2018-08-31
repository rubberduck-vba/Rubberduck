using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteGlobalInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult()
        {
            const string inputCode =
                @"Global var1 As Integer";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult_MultipleGlobals()
        {
            const string inputCode =
                @"Global var1 As Integer
Global var2 As String";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public var1 As Integer";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult_SomeConstantsUsed()
        {
            const string inputCode =
                @"Public var1 As Integer
Global var2 As Date";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ObsoleteGlobal
Global var1 As Integer";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteGlobalInspection";
            var inspection = new ObsoleteGlobalInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
