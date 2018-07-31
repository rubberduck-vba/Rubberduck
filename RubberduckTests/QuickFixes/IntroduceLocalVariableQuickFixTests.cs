using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IntroduceLocalVariableQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_QuickFixWorks()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Collection
    For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            var expectedCode =
                @"Public Sub Foo()
    Dim bar As Collection
    Dim fooBar As Variant
    For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            TestInsertLocalVariableQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_QuickFixWorks_MultipleEnclosingBlocks()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Collection
    With bar
        For Each fooBar In Foo.Items
            fooBar.Whatever
        Next
    End With
End Sub";

            var expectedCode =
                @"Public Sub Foo()
    Dim bar As Collection
    With bar
        Dim fooBar As Variant
        For Each fooBar In Foo.Items
            fooBar.Whatever
        Next
    End With
End Sub";

            TestInsertLocalVariableQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_QuickFixWorks_MultiplePrecedingEndOfLines()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Collection




    For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            var expectedCode =
                @"Public Sub Foo()
    Dim bar As Collection




    Dim fooBar As Variant
    For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            TestInsertLocalVariableQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_QuickFixWorks_PrecedingComment()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Collection 'Comment
    For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            var expectedCode =
                @"Public Sub Foo()
    Dim bar As Collection 'Comment
    Dim fooBar As Variant
    For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            TestInsertLocalVariableQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_QuickFixWorks_LineLabelOnEnclosingStatement()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Collection
l:  For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            var expectedCode =
                @"Public Sub Foo()
    Dim bar As Collection
    Dim fooBar As Variant
l:  For Each fooBar In Foo
        fooBar.Whatever
    Next
End Sub";

            TestInsertLocalVariableQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_QuickFixWorks_StatementSeperators()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Collection : For Each fooBar In Foo : fooBar.Whatever : Next
End Sub";

            var expectedCode =
                @"Public Sub Foo()
    Dim bar As Collection : Dim fooBar As Variant : For Each fooBar In Foo : fooBar.Whatever : Next
End Sub";

            TestInsertLocalVariableQuickFix(expectedCode, inputCode);
        }

        private void TestInsertLocalVariableQuickFix(string expectedCode, string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new UndeclaredVariableInspection(state) { Severity = CodeInspectionSeverity.Warning };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IntroduceLocalVariableQuickFix(state).Fix(inspectionResults.First());
            var actualCode = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actualCode);
        }
    }
}