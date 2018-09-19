using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class DuplicatedAnnotationInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void AnnotationDuplicated_ReturnsResult()
        {
            const string inputCode = @"
Public Sub Bar()
End Sub

'@Obsolete
'@Obsolete
Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void AnnotationDuplicatedTwice_ReturnsSingleResult()
        {
            const string inputCode = @"
Public Sub Bar()
End Sub

'@Obsolete
'@Obsolete
'@Obsolete
Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MultipleAnnotationsDuplicated_ReturnsResult()
        {
            const string inputCode = @"
Public Sub Bar()
End Sub

'@Obsolete
'@Obsolete
'@TestMethod
'@TestMethod
Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void AnnotationNotDuplicated_DoesNotReturnResult()
        {
            const string inputCode = @"
Public Sub Bar()
End Sub

'@Obsolete
Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void AnnotationAllowingMultipleApplicationsDuplicated_DoesNotReturnResult()
        {
            const string inputCode = @"
Public Sub Bar()
End Sub

'@Ignore(Bar)
'@Ignore(Baz)
Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }
    }
}
