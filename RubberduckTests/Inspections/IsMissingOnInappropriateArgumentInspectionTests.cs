using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class IsMissingOnInappropriateArgumentInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ReportsNonVariantOptionalArgument()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ReportsWhenFullyQualified()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print VBA.Information.IsMissing(bar)
End Sub
";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ReportsWhenPartiallyQualified()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print VBA.IsMissing(bar)
End Sub
";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ReportsRequiredVariantArgument()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_DoesNotReportWhenIgnored()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Variant)
'@Ignore IsMissingOnInappropriateArgument
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_DoesNotReportOptionalVariantArgument()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ReportsOptionalVariantArrayArgument()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar() As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ReportsOptionalVariantArgumentWithDefault()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant = 42)
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_DoesNotReportParamArray()
        {
            const string inputCode =
                @"
Public Sub Foo(ParamArray bar() As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_DoesNotReportOptionalVariantArgumentInExpression()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(bar + 1)
End Sub
";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingOnInappropriateArgument_ArgumentAsParameterInExpression()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(Baz(bar))
End Sub

Public Function Baz(arg As Variant) As Variant
End Function
";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        private int ArrangeAndGetInspectionCount(string code)
        {
            var modules = new(string, string, ComponentType)[] 
            {
                ("Module1", code, ComponentType.StandardModule)
            };

            return InspectionResultsForModules(modules, ReferenceLibrary.VBA).Count();
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new IsMissingOnInappropriateArgumentInspection(state);
        }
    }
}
