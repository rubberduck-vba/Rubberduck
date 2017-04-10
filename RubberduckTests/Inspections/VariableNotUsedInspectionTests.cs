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
    public class VariableNotUsedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotUsed_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    Dim var2 As Date
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
    Dim var1 as String
    var1 = ""test""

    Goo var1
End Function

Sub Goo(ByVal arg1 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotUsed_ReturnsResult_MultipleVariables_SomeAssigned()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 as Integer
    var1 = 8

    Dim var2 as String

    Goo var1
End Sub

Sub Goo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    '@Ignore VariableNotUsed
    Dim var1 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableNotUsed_DoesNotReturnsResult_UsedInNameStatement()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    Name ""foo"" As var1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim var1 As String
End Sub";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnassignedVariable_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
Dim var1 As String
End Sub";

            const string expectedCode =
@"Sub Foo()
'@Ignore VariableNotUsed
Dim var1 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableNotUsedInspection(state);
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new VariableNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "VariableNotUsedInspection";
            var inspection = new VariableNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
