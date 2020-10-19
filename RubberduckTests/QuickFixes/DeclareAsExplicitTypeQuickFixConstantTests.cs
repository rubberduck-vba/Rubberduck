using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class DeclareAsExplicitTypeQuickFixConstantTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void Unreferenced_Const()
        {
            var inputCode =
$@"
Private Const MY_CONST = 45.7

Sub Test()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("MY_CONST As Double = 45.7", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ConstantAssignedAConstant()
        {
            var inputCode =
$@"
Private Const OTHER_VALUE As Byte = 200
Private Const MY_CONST = OTHER_VALUE

Sub Test()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("MY_CONST As Byte = OTHER_VALUE", actualCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/3225
        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeclareAsExplicitTypeQuickFixTests))]
        public void ConstUsedAsAParameter()
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

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            StringAssert.Contains("Const BUILD_ON_WEEKDAYS_ONLY As Boolean = True", actualCode);
            StringAssert.Contains("Const DAYS_OF_EFFORT As Long = 2", actualCode);
            StringAssert.Contains("Const WEEKDAY_MULTIPLIER As Double = 1", actualCode);
            StringAssert.Contains("Const WEEKEND_MULTIPLIER As Double = 1.5", actualCode);
            StringAssert.Contains("Const DAILY_RATE As Double = 500", actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new DeclareAsExplicitTypeQuickFix(state, new ParseTreeValueFactory());
        }
    }
}
