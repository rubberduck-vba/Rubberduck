using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
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

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo(ByVal arg1 As Integer) As Boolean
    Let arg1 = 9
End Function";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
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

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
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

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
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

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
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
            var results = GetInspectionResults(vbe.Object);
            Assert.AreEqual(0, results.Count());
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

            var quickFixResult = ApplyIgnoreOnceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult); 
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new AssignedByValParameterInspection(null,null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "AssignedByValParameterInspection";
            var inspection = new AssignedByValParameterInspection(null,null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private void AssertVbaFragmentYieldsExpectedInspectionResultCount(string inputCode, int expectedCount)
        {
            var inspectionResults = GetInspectionResults(inputCode);
            Assert.AreEqual(expectedCount, inspectionResults.Count());
        }

        private string ApplyIgnoreOnceQuickFixToCodeFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe.Object);

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            return GetModuleContent(vbe.Object);
        }

        private string GetModuleContent(IVBE vbe)
        {
            var project = vbe.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            return module.Content();
        }

        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            return GetInspectionResults(vbe.Object);
        }

        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(IVBE vbe)
        {
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new AssignedByValParameterInspection(parser.State,null);
            return inspection.GetInspectionResults();
        }

        private Mock<IVBE> BuildMockVBEStandardModuleForVBAFragment(string inputCode)
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            return builder.BuildFromSingleStandardModule(inputCode, out component);

        }
    }
}
