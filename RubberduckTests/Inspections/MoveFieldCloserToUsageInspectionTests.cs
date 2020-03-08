using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MoveFieldCloseToUsageInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_ReturnsResult()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_MultipleReferenceInDifferentScope()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
    Let bar = ""test""
End Sub
Public Sub For2()
    Let bar = ""test""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnResult_Variable()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim bar As String
    bar = ""test""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_NoReferences()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_ReferenceInPropertyGet()
        {
            const string inputCode =
                @"Private bar As String
Public Property Get Foo() As String
    Foo = bar
End Property";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_ReferenceInPropertyLet()
        {
            const string inputCode =
                @"Private bar As String
Public Property Get Foo() As String
    Foo = ""test""
End Property
Public Property Let Foo(ByVal value As String)
    bar = value
End Property";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_ReferenceInPropertySet()
        {
            const string inputCode =
                @"Private bar As Variant
Public Property Get Foo() As Variant
    Foo = ""test""
End Property
Public Property Set Foo(ByVal value As Variant)
    bar = value
End Property";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore MoveFieldCloserToUsage
Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new MoveFieldCloserToUsageInspection(null);

            Assert.AreEqual(nameof(MoveFieldCloserToUsageInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MoveFieldCloserToUsageInspection(state);
        }
    }
}
