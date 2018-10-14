using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IsMissingOnInappropriateArgumentQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void OptionalStringArgument_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print bar = vbNullString
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionalStringArgumentFullyQualified_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print VBA.Information.IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print bar = vbNullString
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionalStringArgumentPartiallyQualified_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print VBA.IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(Optional bar As String)
    Debug.Print bar = vbNullString
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionalStringArgumentWithDefault_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As Variant = 42)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(Optional bar As Variant = 42)
    Debug.Print bar = 42
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void ParamArray_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(ParamArray bar() As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(ParamArray bar() As Variant)
    Debug.Print LBound(bar) > UBound(bar)
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionalStringArgumentDefaultDoubleQuotes_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar As String = """")
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(Optional bar As String = """")
    Debug.Print bar = vbNullString
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void ReferenceType_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Collection)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As Collection)
    Debug.Print bar Is Nothing
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        private string ArrangeAndApplyQuickFix(string code)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.FirstOrDefault();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new IsMissingOnInappropriateArgumentInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IsMissingOnInappropriateArgumentQuickFix(state).Fix(inspectionResults.First());
                return state.GetRewriter(component).GetText();
            }
        }
    }
}
