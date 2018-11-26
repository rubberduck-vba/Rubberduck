using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class FunctionReturnValueNotUsedInspectionTests
    {
        private IEnumerable<IInspectionResult> GetInspectionResults(string code)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new FunctionReturnValueNotUsedInspection(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_IgnoresUnusedFunction()
        {
            const string code = @"
Public Function Foo() As Long
    Foo = 42
End Function
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_ReturnsResult_ExplicitCallWithoutAssignment()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Call Foo(""Test"")
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_ReturnsResult_CallWithoutAssignment()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Foo ""Test""
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_ReturnsResult_AddressOf()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Bar AddressOf Foo
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_ReturnsResult_NoReturnValueAssignment()
        {
            const string code = @"
Public Function Foo() As Integer
End Function

Public Sub Bar()
    Foo
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_Ignored_DoesNotReturnResult_AddressOf()
        {
            const string code = @"
'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Bar AddressOf Foo
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_MultipleConsecutiveCalls()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Baz()
    Foo Foo(Foo(""Bar""))
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IfStatement()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Baz()
    If Foo(""Test"") Then
    End If
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ForEachStatement()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Sub Bar(ByVal fizz As Boolean)
End Sub

Public Sub Baz()
    For Each Bar In Foo
    Next Bar
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_WhileStatement()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Sub Bar(ByVal fizz As Boolean)
End Sub

Public Sub Baz()
    While Foo
    Wend
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_DoUntilStatement()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Sub Bar(ByVal fizz As Boolean)
End Sub

Public Sub Baz()
    Do Until Foo
    Loop
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ReturnValueAssignment()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Baz()
    TestVal = Foo(""Test"")
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_RecursiveFunction()
        {
            const string code = @"
Public Function Factorial(ByVal n As Long) As Long
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = Factorial(n - 1) * n
    End If
End Function
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ArgumentFunctionCall()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Sub Bar(ByVal fizz As Boolean)
End Sub

Public Sub Baz()
    Bar Foo(""Test"")
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IgnoresBuiltInFunctions()
        {
            const string code = @"
Public Sub Dummy()
    MsgBox ""Test""
    Workbooks.Add
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void GivenInterfaceImplementationMember_ReturnsNoResult()
        {
            const string interfaceCode = @"
Public Function Test() As Integer
End Function
";
            const string implementationCode = @"
Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function
";
            const string callSiteCode = @"
Public Sub Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    Dim result As Integer
    result = testObj.Test
End Sub
";
            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IFoo", ComponentType.ClassModule, interfaceCode)
                .AddComponent("Bar", ComponentType.ClassModule, implementationCode)
                .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                .AddProjectToVbeBuilder().Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new FunctionReturnValueNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_ReturnsResult_InterfaceMember()
        {
            const string interfaceCode = @"
Public Function Test() As Integer
End Function
";
            const string implementationCode = 
                @"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function
";
            const string callSiteCode = @"
Public Sub Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Sub
";
            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IFoo", ComponentType.ClassModule, interfaceCode)
                .AddComponent("Bar", ComponentType.ClassModule, implementationCode)
                .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                .AddProjectToVbeBuilder().Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new FunctionReturnValueNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void InspectionName()
        {
            const string inspectionName = "FunctionReturnValueNotUsedInspection";
            var inspection = new FunctionReturnValueNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
