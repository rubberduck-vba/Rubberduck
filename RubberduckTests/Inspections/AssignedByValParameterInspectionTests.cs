using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

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

        [TestMethod, Ignore]  //Inspections do not find this modification
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_ObjectMethodsCalled()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Collection)
    if arg1.Count = 5 then
        arg1.Add ""Another thing""
    endif
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_QuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";
            const string expectedCode =
@"Public Sub Foo(ByRef arg1 As String)
    Let arg1 = ""test""
End Sub";

            var quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
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

            var quickFixResult = ApplyIgnoreOnceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult); 
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUse()
        {
            //Punt if the user-defined or auto-generated name is already used in the method
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Dim localArg1 as string
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As String)
    Dim localArg1 as string
    Let arg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherSub()
        {
            //Make sure the modified code stays within the specific method under repair
            const string inputCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            const string expectedCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod, Ignore]    //Inspections do not find this modification
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalObjectAssignment()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Collection)
    arg1.Add ""Another thing""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Collection)
Dim localArg1 As Collection
Set localArg1 = arg1
    localArg1.Add ""Another thing""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_FunctionReturn()
        {
            const string inputCode =
@"Private Function MessingWithByValParameters(leaveAlone As Integer, ByVal messWithThis As String) As String
    If leaveAlone > 10 Then
        messWithThis = messWithThis & CStr(leaveAlone)
        messWithThis = Replace(messWithThis, ""OK"", ""yes"")
    End If
    MessingWithByValParameters = messWithThis
End Function
";

            const string expectedCode =
@"Private Function MessingWithByValParameters(leaveAlone As Integer, ByVal messWithThis As String) As String
Dim localMessWithThis As String
localMessWithThis = messWithThis
    If leaveAlone > 10 Then
        localMessWithThis = localMessWithThis & CStr(leaveAlone)
        localMessWithThis = Replace(localMessWithThis, ""OK"", ""yes"")
    End If
    MessingWithByValParameters = localMessWithThis
End Function
";
            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
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

        private string ApplyPassParameterByReferenceQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            inspectionResults.First().QuickFixes.Single(s => s is PassParameterByReferenceQuickFix).Fix();

            return GetModifiedContent(vbe);
        }

        private string ApplyLocalVariableQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            var quickFixBase = inspectionResults.First().QuickFixes.Single(s => s is AssignedByValParameterQuickFix);
            AssignedByValParameterQuickFix assignByValParamQFix = (AssignedByValParameterQuickFix)quickFixBase;
            assignByValParamQFix.ForceFixUsingGeneratedName();
            return GetModifiedContent(vbe);
        }
        private string ApplyIgnoreOnceQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            return GetModifiedContent(vbe);
        }
        private string GetModifiedContent(Mock<IVBE> vbe)
        {
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            return module.Content();
        }
        private System.Collections.Generic.IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            return GetInspectionResults(vbe);
        }
        private System.Collections.Generic.IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(Mock<IVBE> vbe)
        {
            var parser = GetParseCoordinator(vbe);
            var inspection = new AssignedByValParameterInspection(parser.State);
            return inspection.GetInspectionResults();
        }
        private void AssertVbaFragmentYieldsExpectedInspectionResultCount(string inputCode, int expectedCount)
        {
            var inspectionResults = GetInspectionResults(inputCode);
            Assert.AreEqual(expectedCount, inspectionResults.Count());
        }
        private ParseCoordinator GetParseCoordinatorForVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            return GetParseCoordinator(vbe);
        }
        private Mock<IVBE> BuildMockVBEStandardModuleForVBAFragment(string inputCode)
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            return builder.BuildFromSingleStandardModule(inputCode, out component);
            //TODO: removal of the two lines below have no effect on the outcome of any test...remove?
            //var mockHost = new Mock<IHostApplication>();
            //mockHost.SetupAllProperties();
        }
        private ParseCoordinator GetParseCoordinator(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            return parser;
        }
    }
}
