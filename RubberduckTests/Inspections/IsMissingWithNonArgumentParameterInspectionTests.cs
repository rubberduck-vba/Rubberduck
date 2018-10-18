using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class IsMissingWithNonArgumentParameterInspectionTests
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

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
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

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        private int ArrangeAndGetInspectionCount(string code)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .Build();
            var vbe = builder.AddProject(project).Build();


            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new IsMissingWithNonArgumentParameterInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                return inspectionResults.Count();
            }
        }
    }
}
