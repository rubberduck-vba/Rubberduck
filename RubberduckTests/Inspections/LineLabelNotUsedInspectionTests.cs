using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class LineLabelNotUsedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
label1:
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
label1:
    GoTo label1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelUsed_GoToBeforeLabel_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    GoTo label1
label1:
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelUsed_MultipleLabels_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
label1:
    GoTo label1
label2:
    Goto label2
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void LabelNotUsed_ReturnsResult_MultipleLabels_SomeAssigned()
        {
            const string inputCode =
@"Sub Foo()
    GoTo Label1:

label2:

label1:
End Sub

Sub Goo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
//        [TestCategory("Inspections")]
//        public void VariableNotUsed_DoesNotReturnsResult_UsedInNameStatement()
//        {
//            const string inputCode =
//@"Sub Foo()
//    Dim var1 As String
//    Name ""foo"" As var1
//End Sub";

//            IVBComponent component;
//            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
//            var state = MockParser.CreateAndParse(vbe.Object);

//            var inspection = new LineLabelNotUsedInspection(state);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.IsFalse(inspectionResults.Any());
//        }

//        [TestMethod]
//        [TestCategory("Inspections")]
//        public void UnassignedVariable_QuickFixWorks()
//        {
//            const string inputCode =
//@"Sub Foo()
//Dim var1 As String
//End Sub";

//            const string expectedCode =
//@"Sub Foo()
//End Sub";

//            IVBComponent component;
//            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
//            var state = MockParser.CreateAndParse(vbe.Object);

//            var inspection = new LineLabelNotUsedInspection(state);
//            new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

//            var rewriter = state.GetRewriter(component);
//            Assert.AreEqual(expectedCode, rewriter.GetText());
//        }

//        [TestMethod]
//        [TestCategory("Inspections")]
//        public void UnassignedVariable_IgnoreQuickFixWorks()
//        {
//            const string inputCode =
//@"Sub Foo()
//Dim var1 As String
//End Sub";

//            const string expectedCode =
//@"Sub Foo()
//'@Ignore VariableNotUsed
//Dim var1 As String
//End Sub";

//            IVBComponent component;
//            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
//            var state = MockParser.CreateAndParse(vbe.Object);

//            var inspection = new LineLabelNotUsedInspection(state);
//            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspection.GetInspectionResults().First());

//            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
//        }

//        [TestMethod]
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
