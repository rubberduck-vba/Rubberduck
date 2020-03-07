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
    public class IsMissingWithNonArgumentParameterInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_ReportsOptionalVariantArgumentInExpression()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(bar + 1)
End Sub
";

            Assert.AreEqual(1, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_ReportsLocalVariable()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Dim baz As Variant
    Debug.Print IsMissing(baz)
End Sub
";
            Assert.AreEqual(1, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_ReportsFunctionCall()
        {
            const string inputCode =
                @"
Public Sub Foo()
    Debug.Print IsMissing(Bar)
End Sub

Public Function Bar() As Variant
End Function
";
            Assert.AreEqual(1, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_ReportsParameterNameInLiteral()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(""bar"")
End Sub
";
            Assert.AreEqual(1, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_DoesNotReportWhenIgnored()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Variant)
'@Ignore IsMissingWithNonArgumentParameter
    Debug.Print IsMissing(42)
End Sub
";
            Assert.AreEqual(0, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_DoesNotReportOptionalVariantArgument()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(bar)
End Sub
";
            Assert.AreEqual(0, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_DoesNotReportParamArray()
        {
            const string inputCode =
                @"
Public Sub Foo(ParamArray bar() As Variant)
    Debug.Print IsMissing(bar)
End Sub
";
            Assert.AreEqual(0, ArrangeAndGetInspectionCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void IsMissingWithNonArgumentParameter_ReportsArgumentAsParameterInExpression()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant)
    Debug.Print IsMissing(Baz(bar))
End Sub

Public Function Baz(arg As Variant) As Variant
End Function
";
            Assert.AreEqual(1, ArrangeAndGetInspectionCount(inputCode));
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
            return new IsMissingWithNonArgumentParameterInspection(state);
        }
    }
}
