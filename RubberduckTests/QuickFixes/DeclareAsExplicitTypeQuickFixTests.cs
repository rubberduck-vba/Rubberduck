using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class DeclareAsExplicitTypeQuickFixTests : QuickFixTestBase
    {


        [TestCase("var1 = 42", "Integer")]
        [TestCase("var1 = 42.5", "Double")]
        [TestCase("var1 = False", "Boolean")]
        [TestCase(@"var1 = ""StringLiteral""", "String")]
        [TestCase("var1 = #2015-05-15#", "Date")]
        [TestCase("var1 = 45E10", "Double")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void LocalVariable_AssignedUsingLiteralExpresssion(string assignment, string expectedType)
        {
            var inputCode =
$@"Sub Foo()
    Dim var1
    {assignment}
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"As {expectedType}", actualCode);
        }


        [TestCase("var1 = 42", "Long", "Long")]
        [TestCase("var1 = 42.55", "Long", "Double")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void LocalVariable_AssignedUsingLiteralExpresssionAndFunctions(string assignment, string functionType, string expectedType)
        {
            var inputCode =
$@"Sub Foo()
    Dim var1
    {assignment}

    var1 = AssignAValue()
End Sub

Private Function AssignAValue() As {functionType}
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Dim var1 As {expectedType}", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void FixMutipleResults()
        {
            var inputCode =
$@"
Private mTest2
Private mTest

Sub Foo(ByVal arg As String)
    mTest = AssignAValue()
    mTest2 = arg
End Sub

Private Function AssignAValue() As Long
End Function
";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Private mTest2 As String", actualCode);
            StringAssert.Contains($"Private mTest As Long", actualCode);
        }

        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void SetAssignment_FromFunction()
        {
            var inputCode =
$@"
Private mTest

Private mColl As Collection

Sub Fizz()
    Set mTest = AssignACollection()
End Sub

Private Function AssignACollection() As Collection
    Set AssignACollection = mColl
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Private mTest As Collection", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void SetAssignment_FromNew()
        {
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Set mTest = New Collection
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Private mTest As Object", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void SetAssignment_FromOtherInstance()
        {
            var inputCode =
$@"
Private mTest

Private mColl As Collection

Sub Fizz()
    Set mColl = New Collection
    Set mTest = mColl
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Private mTest As Collection", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void UseParameterType_Function()
        {
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    local = AssignAValue(mTest)
End Sub

Private Function AssignAValue(arg As String) As String
    AssignAValue = arg & ""MoreContent""
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Private mTest As String", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void UseParameterType_Sub()
        {
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    AssignAValue(mTest)
End Sub

Private Sub AssignAValue(ByRef arg As String)
    arg = arg & ""MoreContent""
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Private mTest As String", actualCode);
        }

        [TestCase("6, 7, mTest, 9")]
        [TestCase("arg:=mTest, bogey3:=9, bogey2:=7, bogey1:=6")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void UseParameterType_MultipleParameters(string argList)
        {
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    local = AssignAValue(mTest)
    local = AssignAValue2({argList})
End Sub

Private Function AssignAValue(arg As Integer) As String
    AssignAValue = CStr(arg)
End Function

Private Function AssignAValue2(bogey1 As Long, bogey2 As Long, arg As Double, bogey3 As Long) As String
    AssignAValue2 = CStr(arg)
End Function
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Private mTest As Double", actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void UseParameterType_PartOfAParamArray()
        {
            var inputCode =
$@"
Private mTest

Sub Fizz()
    Dim local As String
    Dim anotherLocal As Double
    local = AssignAValue(6, anotherLocal, mTest)
End Sub

Private Function AssignAValue(arg As Integer, ByVal ParamArray args() As Double) As String
    AssignAValue = CStr(arg)
End Function
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Private mTest As Double", actualCode);
        }

        [TestCase("String", "Long", "Variant")]
        [TestCase("String", "String", "String")]
        [TestCase("Date", "Long", "Variant")]
        [TestCase("Date", "Date", "Date")]
        [TestCase("Currency", "Long", "Variant")]
        [TestCase("Currency", "Currency", "Currency")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ModuleVariable_RestrictiveAsTypes(string argType, string functionType, string expected)
        {
            var inputCode =
$@"
Private mTest

Sub Fizz()
    mTest = AssignAValue()
End Sub

Sub Fazz(arg As {argType})
    mTest = arg
End Sub

Private Function AssignAValue() As {functionType}
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains($"Private mTest As {expected}", actualCode);
        }

        [Test]
        [Ignore("Activate once unified expression engine in place")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void BooleanExpression_UsingLiterals()
        {
            var inputCode =
$@"
Private mTest

Sub Test()
    mTest = 3 > 2
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Private mTest As Boolean", actualCode);
        }

        [Test]
        [Ignore("Activate once unified expression engine in place")]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void BooleanNotExpression_UsingFunction()
        {
            var inputCode =
$@"
Private mTest

Sub Test()
    mTest = Not GetBoolean()
End Sub

Private Function GetBoolean() As Boolean
    GetBoolean = true
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Private mTest As Boolean", actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new DeclareAsExplicitTypeQuickFix(state, new ParseTreeValueFactory());
        }
    }
}
