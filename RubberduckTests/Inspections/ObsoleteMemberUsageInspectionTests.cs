using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteMemberUsageInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteMemberUsed_ReturnsResult()
        {
            const string inputCode = @"
'@Obsolete
Public Sub Foo()
End Sub

Public Sub Bar()
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteMemberUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteMemberUsedTwice_ReturnsTwoResults()
        {
            const string inputCode = @"
'@Obsolete
Public Sub Foo()
End Sub

Public Sub Bar()
    Foo
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteMemberUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteMemberUsedOnNonMemberDeclaration_DoesNotReturnResult()
        {
            const string inputCode = @"
'@Obsolete
Public s As String";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteMemberUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }
    }
}
