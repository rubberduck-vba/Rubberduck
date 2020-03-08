using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MultilineParameterInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void MultilineParameter_ReturnsResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultilineParameter_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal Var1 As Integer)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultilineParameter_ReturnsMultipleResults()
        {
            const string inputCode =
                @"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer, _
    ByVal _
    Var2 _
    As _
    Date)
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultilineParameter_ReturnsResults_SomeParams()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer, ByVal Var2 As Date)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultilineParameter_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore MultilineParameter
Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new MultilineParameterInspection(null);

            Assert.AreEqual(nameof(MultilineParameterInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MultilineParameterInspection(state);
        }
    }
}
