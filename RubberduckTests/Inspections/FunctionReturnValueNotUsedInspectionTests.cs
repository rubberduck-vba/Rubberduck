using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class FunctionReturnValueNotUsedInspectionTests
    {
        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_ExplicitCallWithoutAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Call Foo(""Test"")
End Sub";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_CallWithoutAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Foo ""Test""
End Sub";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_AddressOf()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Bar AddressOf Foo
End Sub";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_NoReturnValueAssignment()
        {
            const string inputCode =
@"Public Function Foo() As Integer
End Function
Public Sub Bar()
    Foo
End Sub";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_Ignored_DoesNotReturnResult_AddressOf()
        {
            const string inputCode =
@"'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Bar AddressOf Foo
End Sub";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_MultipleConsecutiveCalls()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    Foo Foo(Foo(""Bar""))
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IfStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    If Foo(""Test"") Then
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ForEachStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    For Each Bar In Foo
    Next Bar
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_WhileStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    While Foo
    Wend
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_DoUntilStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    Do Until Foo
    Loop
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ReturnValueAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    TestVal = Foo(""Test"")
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_RecursiveFunction()
        {
            const string inputCode =
@"Public Function Factorial(ByVal n As Long) As Long
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = Factorial(n - 1) * n
    End If
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ArgumentFunctionCall()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    Bar Foo(""Test"")
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IgnoresBuiltInFunctions()
        {
            const string inputCode =
@"Public Sub Dummy()
    MsgBox ""Test""
    Workbooks.Add
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void GivenInterfaceImplementationMember_ReturnsNoResult()
        {
            const string interfaceCode =
@"Public Function Test() As Integer
End Function";

            const string implementationCode =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Sub Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    Dim result As Integer
    result = testObj.Test
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                             .AddComponent("IFoo", ComponentType.ClassModule, interfaceCode)
                             .AddComponent("Bar", ComponentType.ClassModule, implementationCode)
                             .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                             .MockVbeBuilder().Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_InterfaceMember()
        {
            const string interfaceCode =
@"Public Function Test() As Integer
End Function";

            const string implementationCode =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Sub Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                                        .AddComponent("IFoo", ComponentType.ClassModule, interfaceCode)
                                        .AddComponent("Bar", ComponentType.ClassModule, implementationCode)
                                        .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                                        .MockVbeBuilder().Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
    If True Then
        Foo = _
        True
    Else
        Foo = False
    End If
End Function";

            const string expectedCode =
@"Public Sub Foo(ByVal bar As String)
    If True Then
        
    Else
        
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface_ManyBodyStatements()
        {
            const string inputCode =
@"Function foo(ByRef fizz As Boolean) As Boolean
    fizz = True
    goo
label1:
    foo = fizz
End Function

Sub goo()
End Sub";

            const string expectedCode =
@"Sub foo(ByRef fizz As Boolean)
    fizz = True
    goo
label1:
    
End Sub

Sub goo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks_Interface()
        {
            const string inputInterfaceCode =
@"Public Function Test() As Integer
End Function";

            const string expectedInterfaceCode =
@"Public Sub Test()
End Sub";

            const string inputImplementationCode1 =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string inputImplementationCode2 =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Function Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                             .AddComponent("IFoo", ComponentType.ClassModule, inputInterfaceCode)
                             .AddComponent("Bar", ComponentType.ClassModule, inputImplementationCode1)
                             .AddComponent("Bar2", ComponentType.ClassModule, inputImplementationCode2)
                             .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                             .MockVbeBuilder().Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());

            var component = vbe.Object.VBProjects[0].VBComponents[0];
            Assert.AreEqual(expectedInterfaceCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            const string expectedCode =
@"'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new FunctionReturnValueNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "FunctionReturnValueNotUsedInspection";
            var inspection = new FunctionReturnValueNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
