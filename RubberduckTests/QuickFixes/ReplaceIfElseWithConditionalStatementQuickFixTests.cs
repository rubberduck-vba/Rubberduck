using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceIfElseWithConditionalStatementQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void Simple()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new BooleanAssignedInIfElseInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ComplexCondition()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Or False And False Xor True Then
        d = True
    Else
        d = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True Or False And False Xor True
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new BooleanAssignedInIfElseInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void InvertedCondition()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = False
    Else
        d = True
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = Not (True)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new BooleanAssignedInIfElseInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void QualifiedName()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Fizz.Buzz = True
    Else
        Fizz.Buzz = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Fizz.Buzz = True
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new BooleanAssignedInIfElseInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ReplaceIfElseWithConditionalStatementQuickFix();
        }
    }
}
