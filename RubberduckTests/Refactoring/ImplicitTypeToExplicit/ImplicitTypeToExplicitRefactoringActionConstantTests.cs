using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.ImplicitTypeToExplicit
{
    [TestFixture]
    public class ImplicitTypeToExplicitRefactoringActionConstantTests : ImplicitTypeToExplicitRefactoringActionTestsBase
    {
        [TestCase("45.7", "Double")]
        [TestCase("45", "Long")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void Unreferenced_Const(string literalValue, string expectedType)
        {
            var targetName = "MY_CONST";
            var inputCode =
$@"
Private Const MY_CONST = {literalValue}
";
            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"Const {targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConstantAssignedAConstant()
        {
            var targetName = "MY_CONST";
            var expectedType = "Byte";
            var inputCode =
$@"
Private Const OTHER_VALUE As {expectedType} = 200
Private Const {targetName} = OTHER_VALUE
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"Const {targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConstantUsedAsValueParameter()
        {
            var targetName = "MY_CONST";
            var expectedType = "Double";
            var inputCode =
$@"
Private Const {targetName} = 55

Sub Test()
    Fizz = {targetName}
End Sub

Public Property Let Fizz(ByVal RHS As Double)
End Property
";

            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"Const {targetName} As {expectedType}", refactoredCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/3225
        [TestCase("BUILD_ON_WEEKDAYS_ONLY", "Boolean")]
        [TestCase("DAYS_OF_EFFORT", "Long")]
        [TestCase("WEEKDAY_MULTIPLIER", "Double")]
        [TestCase("WEEKEND_MULTIPLIER", "Double")]
        [TestCase("DAILY_RATE", "Double")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConstUsedAsAParameter(string targetName, string expectedType)
        {
            var inputCode =
$@"
Option Explicit

Const BUILD_ON_WEEKDAYS_ONLY = True 'Constant is only ever used as a Boolean
Const DAYS_OF_EFFORT = 2            'Constant is only ever used as a Long
Const WEEKDAY_MULTIPLIER = 1        'Constant is only ever used as a Double
Const WEEKEND_MULTIPLIER = 1.5      'Constant is only ever used as a Double
Const DAILY_RATE = 500              'Constant is only ever used as a Double

Sub test()
  PrintProjectCost True
End Sub

Sub PrintProjectCost(Optional BuildOnWeekDays As Boolean = BUILD_ON_WEEKDAYS_ONLY)

  If BuildOnWeekDays Then
    Debug.Print ProjectCost(DAYS_OF_EFFORT, DAILY_RATE, WEEKDAY_MULTIPLIER)
  Else
    Debug.Print ProjectCost(DAYS_OF_EFFORT, DAILY_RATE, WEEKEND_MULTIPLIER)
  End If

End Sub

Private Function ProjectCost(DaysWorked As Long, DailyRate As Double, WeekendMultiplier As Double) As Double
  ProjectCost = DaysWorked * DailyRate * WeekendMultiplier
End Function
";
            var refactoredCode = RefactoredCode(inputCode,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"Const {targetName} As {expectedType}", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void Constant_IsParameterOfExternalFunction()
        {
            var targetName = "MY_CONST";
            var expectedType = "Double";
            var inputCode =
$@"

Private Const MY_CONST = 6

Private Sub Test()
    Fizz(MY_CONST)
End Sub
";
            var assigningModuleCode =
$@"
Public Sub Fizz(arg As Double)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("AssigningModule", assigningModuleCode));
            var refactoredCode = RefactoredCode(vbe.Object,
                state => TestModel(state, NameAndDeclarationTypeTuple(targetName), (model) => model));

            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase("Property", "Get ")]
        [TestCase("Function", "")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConstantTypedByFunctionType(string procedureType, string getToken)
        {
            var targetName = "MY_CONSTANT";
            var expectedType = "Double";
            var inputCode =
$@"
Public Const MY_CONSTANT = 50

Public {procedureType} {getToken}AValue() As Double
    AValue = MY_CONSTANT
End {procedureType}
";
            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        [TestCase("5 & 5", "String")]
        [TestCase("Null & Null", "Variant")]
        [TestCase(@"Null & ""Test""", "String")]
        [TestCase("5 & Null", "String")]
        [Category("Refactorings")]
        [Category(nameof(ImplicitTypeToExplicitRefactoringAction))]
        public void ConstantTypedByConcatOp(string expression, string expectedType)
        {
            var targetName = "MY_CONSTANT";
            var inputCode =
$@"
Public Const MY_CONSTANT = {expression}

";
            var refactoredCode = RefactoredCode(inputCode, NameAndDeclarationTypeTuple(targetName));
            StringAssert.Contains($"{targetName} As {expectedType}", refactoredCode);
        }

        private static (string, DeclarationType) NameAndDeclarationTypeTuple(string name)
            => (name, DeclarationType.Constant);
    }
}
