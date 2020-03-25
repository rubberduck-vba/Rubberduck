using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveStepOneQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void StepOne_QuickFixWorks_Remove()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 5
    Next
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new StepOneIsRedundantInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void StepOne_QuickFixWorks_NestedLoops()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 5
        For value = 0 To 5 Step 1
        Next
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 5
        For value = 0 To 5
        Next
    Next
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new StepOneIsRedundantInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveStepOneQuickFix();
        }
    }
}
