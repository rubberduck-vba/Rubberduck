using System;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnassignedVariableUsageQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UnassignedVariableUsageInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var (actual, inspectionDescription) =
                ApplyQuickFixToFirstInspectionResultWithInspectionResultDescription(
                    inputCode,
                    state => new UnassignedVariableUsageInspection(state));
            var expected = string.Format(expectedCode, inspectionDescription);
            Assert.AreEqual(expected, actual);
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

            var (actual, inspectionDescription) =
                ApplyQuickFixToFirstInspectionResultWithInspectionResultDescription(
                    inputCode,
                    state => new UnassignedVariableUsageInspection(state));
            var expected = string.Format(expectedCode, inspectionDescription);
            Assert.AreEqual(expected, actual);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveUnassignedVariableUsageQuickFix();
        }

        private (string code, string inspectionResultDescription)
            ApplyQuickFixToFirstInspectionResultWithInspectionResultDescription(string inputCode,
                Func<RubberduckParserState, IInspection> inspectionFactory)
        {
            var vbe = TestVbe(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var resultToFix = inspectionResults.First();
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var quickFix = QuickFix(state);

                quickFix.Fix(resultToFix, rewriteSession);

                var code = rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();
                var inspectionDescription = resultToFix.Description;

                return (code, inspectionDescription);
            }
        }
    }
}
