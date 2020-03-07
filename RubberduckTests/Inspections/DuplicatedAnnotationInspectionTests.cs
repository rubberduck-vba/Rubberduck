using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class DuplicatedAnnotationInspectionTests : InspectionTestsBase
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

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
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

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
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

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
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

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
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

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new DuplicatedAnnotationInspection(null);

            Assert.AreEqual(nameof(DuplicatedAnnotationInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new DuplicatedAnnotationInspection(state);
        }
    }
}
