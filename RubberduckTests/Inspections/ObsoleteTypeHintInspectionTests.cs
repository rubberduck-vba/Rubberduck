using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteTypeHintInspectionTests : InspectionTestsBase
    {
        [TestCase("Public Foo&")]
        [TestCase("Public Foo%")]
        [TestCase("Public Foo#")]
        [TestCase("Public Foo!")]
        [TestCase("Public Foo@")]
        [TestCase("Public Foo$")]
        [TestCase("Public Function Foo$(ByVal bar As Boolean)\r\nEnd Function")]
        [TestCase("Public Property Get Foo$(ByVal bar As Boolean)\r\nEnd Property")]
        [Category("Inspections")]
        public void ObsoleteTypeHint_VariousTypeHints_ReturnsResult(string inputCode)
        {
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteTypeHint_ParameterReturnsResult()
        {
            const string inputCode =
                @"Public Function Foo(ByVal bar$) As Boolean
End Function";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteTypeHint_VariableReturnsResult()
        {
            const string inputCode =
                @"Public Function Foo() As Boolean
    Dim buzz$
    Foo = True
End Function";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteTypeHint_ConstantReturnsResult()
        {
            const string inputCode =
                @"Public Function Foo() As Boolean
    Const buzz$ = 0
    Foo = True
End Function";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteTypeHint_StringValueDoesNotReturnsResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim bar As String
    bar = ""Public baz$""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteTypeHint_FieldsReturnMultipleResults()
        {
            const string inputCode =
                @"Public Foo$
Public Bar$";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteTypeHint_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ObsoleteTypeHint
Public Function Foo$(ByVal bar As Boolean)
End Function";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteTypeHintInspection(null);

            Assert.AreEqual(nameof(ObsoleteTypeHintInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteTypeHintInspection(state);
        }
    }
}
