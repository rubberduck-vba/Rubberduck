using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class DeclareAsExplicitTypeQuickFixParameterTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void Parameter_NoAssignmentsToEvaluate()
        {
            const string inputCode =
                @"Sub Foo(arg1)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg1 As Variant)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void SubNameContainsParameterName()
        {
            const string inputCode =
                @"Sub Foo(Foo)
End Sub";

            const string expectedCode =
                @"Sub Foo(Foo As Variant)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ParameterTypeUndefined_UsedAsArgument()
        {
            var inputCode =
$@"

Private Sub FirstSub(arg)
    AnotherSub arg
End Sub

Private Sub AnotherSub(arg As String)
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"FirstSub(arg As String)", actualCode);
        }

        [TestCase("Optional ByVal Fizz = False", "Optional ByVal Fizz As Boolean = False")]
        [TestCase("Optional ByVal Fizz = #2015-05-15#", "Optional ByVal Fizz As Date = #2015-05-15#")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ParameterWithDefaultValue_Literal(string argList, string expectedArgList)
        {
            var inputCode =
$@"Sub Test({argList})
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains(expectedArgList, actualCode);
        }


        [TestCase("Optional ByVal Fizz = MY_CONST", "Optional ByVal Fizz As Double = MY_CONST")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ParameterWithDefaultValue_Const(string argList, string expectedArgList)
        {
            var inputCode =
$@"
Private Const MY_CONST As Double = 45.7

Sub Test({argList})
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains(expectedArgList, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ParameterAssignedWithinProcedure()
        {
            var inputCode =
$@"

Private aNumber As Double

Sub Test(ByVal multiplier As Long,  fizz)
    aNumber = mulitplier * .5
    fizz = aNumber
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("fizz As Double", actualCode);
        }

        [Test]
        [Ignore("Activate once unified expression engine in place")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ParameterWithDefaultValue_Expression()
        {
            var inputCode =
$@"
Sub Test(Optional ByVal fizz = 7 + 10)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Optional ByVal fizz As Integer = 7 + 10", actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new DeclareAsExplicitTypeQuickFix(state, new ParseTreeValueFactory());
        }
    }
}
