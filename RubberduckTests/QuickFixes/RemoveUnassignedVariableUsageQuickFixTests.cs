using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnassignedVariableUsageQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariableUsage_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveUnassignedVariableUsageQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        // See https://github.com/rubberduck-vba/Rubberduck/issues/3636
        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariableUsage_QuickFixWorksWithBlock()
        {
            const string inputCode =
                @"Sub test()
    Dim wb As Workbook
    With wb
        Debug.Print .Name
        Debug.Print .Name
        Debug.Print .Name
    End With
End Sub";

            const string expectedCode =
                @"Sub test()
    Dim wb As Workbook
    'TODO - {0}
'    With wb
'        Debug.Print .Name
'        Debug.Print .Name
'        Debug.Print .Name
'    End With
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResult = inspection.GetInspectionResults(CancellationToken.None).First();
                var expected = string.Format(expectedCode, inspectionResult.Description);

                new RemoveUnassignedVariableUsageQuickFix(state).Fix(inspectionResult);
                var actual = state.GetRewriter(component).GetText();
                Assert.AreEqual(expected, actual);
            }
        }

        [Test]
        [Ignore("Passes when run individually, does not pass when all tests are run.")]
        [Category("QuickFixes")]
        public void UnassignedVariableUsage_QuickFixWorksNestedWithBlock()
        {
            const string inputCode =
                @"Sub test()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    With wb
        Debug.Print .Name
        With ws
            Debug.Print .Name
            Debug.Print .Name
            Debug.Print .Name
        End With
    End With
End Sub";

            const string expectedCode =
                @"Sub test()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    With wb
        Debug.Print .Name
        'TODO - {0}
'        With ws
'            Debug.Print .Name
'            Debug.Print .Name
'            Debug.Print .Name
'        End With
    End With
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResult = inspection.GetInspectionResults(CancellationToken.None).First();
                var expected = string.Format(expectedCode, inspectionResult.Description);

                new RemoveUnassignedVariableUsageQuickFix(state).Fix(inspectionResult);
                var actual = state.GetRewriter(component).GetText();
                Assert.AreEqual(expected, actual);
            }
        }
    }
}
