using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class AssignedByValParameterInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Sub()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo(ByVal arg1 As Integer) As Boolean
    Let arg1 = 9
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_MultipleParams()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
    Let arg1 = ""test""
    Let arg2 = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_Ignored_DoesNotReturnResult_Sub()
        {
            const string inputCode =
@"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_SomeAssignedByValParams()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
    Let arg1 = ""test""
    
    Dim var1 As Integer
    var1 = arg2
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_NoResultForLeftHandSideMemberAssignment()
        {
            var class1 = @"
Option Explicit
Private mSomething As Long
Public Property Get Something() As Long
    Something = mSomething
End Property
Public Property Let Something(ByVal value As Long)
    mSomething = value
End Property
";
            var caller = @"
Option Explicit
Private Sub DoSomething(ByVal foo As Class1)
    foo.Something = 42
End Sub
";
            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, caller)
                .MockVbeBuilder()
                .Build();
            
            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new AssignedByValParameterInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "AssignedByValParameterInspection";
            var inspection = new AssignedByValParameterInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
