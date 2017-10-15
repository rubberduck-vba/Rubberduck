using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MultipleDeclarationsInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_ReturnsResult_Variables()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_ReturnsResult_Constants()
        {
            const string inputCode =
@"Public Sub Foo()
    Const var1 As Integer = 9, var2 As String = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_ReturnsResult_StaticVariables()
        {
            const string inputCode =
@"Public Sub Foo()
    Static var1 As Integer, var2 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_ReturnsResult_MultipleDeclarations()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer, var2 As String
    Dim var3 As Boolean, var4 As Date
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_ReturnsResult_SomeDeclarationsSeparate()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer, var2 As String
    Dim var3 As Boolean
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_Ignore_DoesNotReturnResult_Variables()
        {
            const string inputCode =
@"Public Sub Foo()
    '@Ignore MultipleDeclarations
    Dim var1 As Integer, var2 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_QuickFixWorks_Variables()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
@"Public Sub Foo()
Dim var1 As Integer
Dim var2 As String

End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_QuickFixWorks_Constants()
        {
            const string inputCode =
@"Public Sub Foo()
    Const var1 As Integer = 9, var2 As String = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo()
Const var1 As Integer = 9
Const var2 As String = ""test""

End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_QuickFixWorks_StaticVariables()
        {
            const string inputCode =
@"Public Sub Foo()
    Static var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
@"Public Sub Foo()
Static var1 As Integer
Static var2 As String

End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultipleDeclarations_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
@"Public Sub Foo()
'@Ignore MultipleDeclarations
    Dim var1 As Integer, var2 As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultipleDeclarationsInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new MultipleDeclarationsInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "MultipleDeclarationsInspection";
            var inspection = new MultipleDeclarationsInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
