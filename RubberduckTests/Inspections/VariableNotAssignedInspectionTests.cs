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
    public class VariableNotAssignedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotAssigned_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    Dim var2 As Date
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_DoesNotReturnResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
    Dim var1 as String
    var1 = ""test""
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_ReturnsResult_MultipleVariables_SomeAssigned()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 as Integer
    var1 = 8

    Dim var2 as String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotAssigned_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
'@Ignore VariableNotAssigned
Dim var1 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim var1 as Integer
End Sub";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            new RemoveUnassignedIdentifierQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void UnassignedVariable_VariableOnMultipleLines_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim _
var1 _
as _
Integer
End Sub";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            new RemoveUnassignedIdentifierQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void UnassignedVariable_MultipleVariablesOnSingleLine_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim var1 As Integer, var2 As Boolean
End Sub";

            // note the extra space after "Integer"--the VBE will remove it
            const string expectedCode =
@"Sub Foo()
Dim var1 As Integer
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            new RemoveUnassignedIdentifierQuickFix(state).Fix(
                inspection.GetInspectionResults().Single(s => s.Target.IdentifierName == "var2"));

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void UnassignedVariable_MultipleVariablesOnMultipleLines_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim var1 As Integer, _
var2 As Boolean
End Sub";

            const string expectedCode =
@"Sub Foo()
Dim var1 As Integer
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            new RemoveUnassignedIdentifierQuickFix(state).Fix(
                inspection.GetInspectionResults().Single(s => s.Target.IdentifierName == "var2"));

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim var1 as Integer
End Sub";

            const string expectedCode =
@"Sub Foo()
'@Ignore VariableNotAssigned
Dim var1 as Integer
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotAssignedInspection(state);
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new VariableNotAssignedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "VariableNotAssignedInspection";
            var inspection = new VariableNotAssignedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
