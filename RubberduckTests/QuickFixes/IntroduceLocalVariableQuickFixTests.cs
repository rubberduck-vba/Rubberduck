using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IntroduceLocalVariableQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UndeclaredVariableInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UndeclaredVariableInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UndeclaredVariableInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UndeclaredVariableInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UndeclaredVariableInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UndeclaredVariableInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new IntroduceLocalVariableQuickFix();
        }
    }
}