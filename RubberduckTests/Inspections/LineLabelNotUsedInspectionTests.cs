using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class LineLabelNotUsedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_EdgeCaseIssue3226()
        {
            const string inputCode = @"
Sub foo()

100           'line-number
200:          'line-number with instruction spearator
300 Beep      'Line-number and instruction
400: Beep     'Line-number with instruction separator and instruction

bar:         'line-label
buzz: Beep   'Line-label and instruction
50 fizz:     'line-number and line-label
10 foo: Beep 'Line number and line-label (that matches procedure name) and instruction
20 boo: Beep 'Line number and line-label and instruction

End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(5, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
label1:
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_ReturnsResult_MultipleLabels()
        {
            const string inputCode =
@"Sub Foo()
label1:
label2:
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenGoToStatement_LabelUsed_YieldsNoResult()
        {
            const string inputCode =
@"Sub Foo()
label1:
    GoTo label1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenGoToStatement_GoToBeforeLabel_LabelUsed_YieldsNoResult()
        {
            const string inputCode =
@"Sub Foo()
    GoTo label1
label1:
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenMultipleGoToStatements_BothLabelsUsed_YieldsNoResult()
        {
            const string inputCode =
@"Sub Foo()
label1:
    GoTo label1
label2:
    Goto label2
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_ReturnsResult_WithUsedLabelThatDoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    GoTo Label1:

label2:

label1:
End Sub

Sub Goo(ByVal arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
'@Ignore label1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
label1:
End Sub";

            const string expectedCode =
@"Sub Foo()

End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_QuickFixWorks_MultipleLabels()
        {
            const string inputCode =
@"Sub Foo()
label1:
dim var1 as variant
label2:
goto label1:
End Sub";

            const string expectedCode =
@"Sub Foo()
label1:
dim var1 as variant

goto label1:
End Sub"; ;

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]

        [TestCategory("Inspections")]
        public void LabelNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
label1:
End Sub";

            const string expectedCode =
@"Sub Foo()
'@Ignore LineLabelNotUsed
label1:
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new LineLabelNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "LineLabelNotUsedInspection";
            var inspection = new LineLabelNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
