using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;

namespace RubberduckTests.Refactoring.ImplicitTypeToExplicit
{
    [TestFixture]
    public class ImplicitTypeToExplicitRefactoringActionParameterTests : ImplicitTypeToExplicitRefactoringActionTestsBase
    {
        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void Parameter_NoAssignmentsToEvaluate()
        {
            var targetName = "arg1";
            var expectedType = "Variant";

            var inputCode =
@"Sub Foo(arg1)
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterUsedAsValueParameter()
        {
            var targetName = "arg";
            var expectedType = "Double";
            var inputCode =
$@"

Sub Test({targetName})
    Fizz = {targetName}
End Sub

Public Property Let Fizz(ByVal RHS As Double)
End Property
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void SubNameContainsParameterName()
        {

            var targetName = "Foo";
            var expectedType = "Variant";
            var inputCode =
@"Sub Foo(Foo)
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterTypeUndefined_UsedAsArgument()
        {
            var targetName = "arg";
            var expectedType = "String";
            var inputCode =
@"
Private Sub FirstSub(arg)
    AnotherSub arg
End Sub

Private Sub AnotherSub(arg2 As String)
End Sub
";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("Optional ByVal Fizz = False", "Boolean")]
        [TestCase("Optional ByVal Fizz = #2015-05-15#", "Date")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterWithDefaultValue_Literal(string argList, string expectedType)
        {
            var targetName = "Fizz";
            var inputCode =
$@"Sub Test({argList})
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("Optional ByVal fizz = MY_CONST", "Optional ByVal fizz As Double = MY_CONST")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterWithDefaultValue_Const(string argList, string expectedArgList)
        {
            var targetName = "fizz";
            var expectedType = "Double";
            var inputCode =
$@"
Private Const MY_CONST As Double = 45.7

Sub Test({argList})
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("ByRef fizz = 3.14", "ByRef fizz As Double = 3.14")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterDefaultValueOverridesAssignment(string argList, string expectedArgList)
        {
            var targetName = "fizz";
            var expectedType = "Double";
            var inputCode =
$@"

Sub Test({argList})
    Dim local As Long
    local = 6
    fizz = local
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterDefaultValueAndExpressionAssignment()
        {
            var targetName = "fizz";
            //TODO: Once an expression evaluation capability is added, this should resolve to Double
            var expectedType = "Variant";
            var inputCode =
$@"

Sub Test(ByRef fizz = 5)
    fizz = 3.14 * 5
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterAssignedWithinProcedure()
        {
            var targetName = "fizz";
            var expectedType = "Double";
            var inputCode =
@"
Private aNumber As Double

Sub Test(ByVal multiplier As Long,  fizz)
    aNumber = multiplier * .5
    fizz = aNumber
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        //TODO: Once an expression evaluation capability is added, this should resolve to Long
        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterWithDefaultValue_Expression()
        {
            var targetName = "fizz";
            var expectedType = "Variant";
            var inputCode =
$@"
Sub Test(Optional ByVal fizz = 7 + 10)
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        //TODO: Once an expression evaluation capability is added, this should resolve to Double
        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterWithDefaultValue_ExpressionAndAssignment_Indeterminant()
        {
            var targetName = "fizz";
            var expectedType = "Variant";
            var inputCode =
$@"
Private mTest As Long

Sub Test(Optional ByRef fizz = 55.6 / 2.2)
    fizz = mTest
End Sub
";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterTypedByFunctionType()
        {
            var targetName = "arg";
            var expectedType = "Double";
            var inputCode =
$@"
Public Function ReturnSecondParameter(ByVal arg1 As Long, arg) As Double
    ReturnSecondParameter = arg
End Function
";
            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5646        
        [TestCase("argArray()", "argArray() As Variant")]
        [TestCase("arg1 As Long, arg2 As String, argArray()", "arg1 As Long, arg2 As String, argArray() As Variant")]
        [TestCase("arg1 As Long, argArray(), arg2 As String", "arg1 As Long, argArray() As Variant, arg2 As String")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void Parameter_Arrays(string argList, string expectedArgList)
        {
            var targetName = "argArray";

            var inputCode =
$@"Sub Foo({argList})
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains(expectedArgList, refactoredCode);
        }

        [TestCase("5 & 5", "String")]
        [TestCase("Null & Null", "Variant")]
        [TestCase(@"Null & ""Test""", "String")]
        [TestCase(@"5 & Null", "String")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterWithDefaultValue_ConcatExpression(string expression, string expected)
        {
            var targetName = "fizz";
            var inputCode =
$@"
Sub Test(Optional ByVal fizz = {expression})
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expected}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ParameterAssignedWithinProcedure_ConcatExpression()
        {
            var targetName = "fizz";
            var expectedType = "String";
            var inputCode =
@"
Sub Test(ByVal countSuffix As Long,  fizz)
    fizz = ""Test"" & countSuffix
End Sub";

            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        (string, DeclarationType) NameAndDeclarationTypeTuple(string name)
            => (name, DeclarationType.Parameter);
    }
}