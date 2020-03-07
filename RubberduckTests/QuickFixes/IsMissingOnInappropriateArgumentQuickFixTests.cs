using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IsMissingOnInappropriateArgumentQuickFixTests : QuickFixTestBase
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
        public void OptionalArray_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(Optional bar() As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(Optional bar() As Variant)
    Debug.Print LBound(bar) > UBound(bar)
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalArray_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar() As String)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar() As String)
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
        public void NonOptionalReferenceType_QuickFixWorks()
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

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalObject_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Object)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As Object)
    Debug.Print bar Is Nothing
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalNumeric_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Long)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As Long)
    Debug.Print bar = 0
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalVariant_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Variant)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As Variant)
    Debug.Print IsEmpty(bar)
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalImplicitVariant_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar)
    Debug.Print IsEmpty(bar)
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalString_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As String)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As String)
    Debug.Print bar = vbNullString
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalDate_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Date)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As Date)
    Debug.Print bar = CDate(0)
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalBoolean_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As Boolean)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As Boolean)
    Debug.Print bar = False
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalTypeHinted_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo(bar&)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar&)
    Debug.Print bar = 0
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalEnumeration_QuickFixPicksMember()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As VbVarType)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As VbVarType)
    Debug.Print bar = vbEmpty
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonOptionalEnumerationNoDefault_QuickFixPicksZero()
        {
            const string inputCode =
                @"
Public Sub Foo(bar As VbStrConv)
    Debug.Print IsMissing(bar)
End Sub
";

            const string expected =
                @"
Public Sub Foo(bar As VbStrConv)
    Debug.Print bar = 0
End Sub
";

            var actual = ArrangeAndApplyQuickFix(inputCode);
            Assert.AreEqual(expected, actual);
        }

        private string ArrangeAndApplyQuickFix(string code)
        {
            return ApplyQuickFixToFirstInspectionResult(code, state => new IsMissingOnInappropriateArgumentInspection(state));
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new IsMissingOnInappropriateArgumentQuickFix(state);
        }

        protected override IVBE TestVbe(string code, out IVBComponent component)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference(ReferenceLibrary.VBA)
                .Build();
            var vbe = builder.AddProject(project).Build();
            component = project.Object.VBComponents.First();
            return vbe.Object;
        }
    }
}
