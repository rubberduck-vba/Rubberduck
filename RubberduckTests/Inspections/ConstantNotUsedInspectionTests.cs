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
    public class ConstantNotUsedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_ReturnsResult_MultipleConsts()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
    Const const2 As String = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_ReturnsResult_UnusedConstant()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
    Goo const1

    Const const2 As String = ""test""
End Sub

Public Sub Goo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
    Goo const1
End Sub

Public Sub Goo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_IgnoreModule_All_YieldsNoResult()
        {
            const string inputCode =
@"'@IgnoreModule

Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_IgnoreModule_AnnotationName_YieldsNoResult()
        {
            const string inputCode =
@"'@IgnoreModule ConstantNotUsed

Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_IgnoreModule_OtherAnnotationName_YieldsResults()
        {
            const string inputCode =
@"'@IgnoreModule VariableNotUsed

Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsTrue(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo()
    '@Ignore ConstantNotUsed
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
Const const1 As Integer = 9
End Sub";

            const string expectedCode =
@"Public Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new RemoveUnusedDeclarationQuickFix(state).Fix(inspectionResults.First());

            var rewriter = state.GetRewriter(component);
            var rewrittenCode = rewriter.GetText();
            Assert.AreEqual(expectedCode, rewrittenCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ConstantNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
@"Public Sub Foo()
'@Ignore ConstantNotUsed
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ConstantNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ConstantNotUsedInspection";
            var inspection = new ConstantNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
