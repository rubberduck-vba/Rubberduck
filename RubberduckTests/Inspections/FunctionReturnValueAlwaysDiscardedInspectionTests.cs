using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class FunctionReturnValueAlwaysDiscardedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void IgnoresUnusedFunction()
        {
            const string code = @"
Public Function Foo() As Long
    Foo = 42
End Function
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void ExplicitCallWithoutAssignment_ReturnsResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Call Foo(""Test"")
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void CallWithoutAssignment_ReturnsResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Foo ""Test""
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void ReturnValueAssignment_DoesNotReturnResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Baz()
    TestVal = Foo(""Test"")
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void CallWithoutAssignmentAndUseOfReturnValue_NoResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Foo ""Test""
End Sub

Public Sub Baz()
    Dim var As Integer
    var = Foo(""Test"")
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void TwoUses_OneResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Foo ""Test""
End Sub

Public Sub Baz()
    Foo ""Test""
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void AddressOfAndCallWithoutAssignment_NoResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Foo ""Test""
End Sub

Public Sub Baz()
    Bar AddressOf Foo
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void AddressOfAlone_NoResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Bar()
    Bar AddressOf Foo
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void NoReturnValueAssignment_ReturnsResult()
        {
            const string code = @"
Public Function Foo() As Integer
End Function

Public Sub Bar()
    Foo
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void Ignored_DoesNotReturnResult()
        {
            const string code = @"
'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function Foo() As Integer
End Function

Public Sub Bar()
    Foo
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void DoesNotReturnResult_MultipleConsecutiveCalls()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Baz()
    Foo Foo(Foo(""Bar""))
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void IfStatement_DoesNotReturnResult()
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void ForEachStatement_DoesNotReturnResult()
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void WhileStatement_DoesNotReturnResult()
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void DoUntilStatement_DoesNotReturnResult()
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void WithStatement_DoesNotReturnResult()
        {
            const string code = @"
Public Function Foo() As Object
End Function

Public Sub Baz()
    With Foo
        'Do Whatever
    End With
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void RecursiveFunction_DoesNotReturnResult()
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void ArgumentFunctionCall_DoesNotReturnResult()
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void OutputListFunctionCall_DoesNotReturnResult()
        {
            const string code = @"
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function

Public Sub Baz()
    Debug.Print Foo(""Test"")
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void IgnoresBuiltInFunctions_DoesNotReturnResult()
        {
            const string code = @"
Public Sub Dummy()
    MsgBox ""Test""
    Workbooks.Add
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(code).Count());
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
            var modules = new (string, string, ComponentType)[] 
            {
                ("IFoo", interfaceCode, ComponentType.ClassModule),
                ("Bar", implementationCode, ComponentType.ClassModule),
                ("TestModule", callSiteCode, ComponentType.StandardModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }


        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void ChainedMemberAccess_ReturnsNoResult()
        {
            const string inputCode = @"
Public Function GetIt(x As Long) As Object
End Function

Public Sub Baz()
    GetIt(1).Select
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }


        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void ChainedParameterlessMemberAccess_ReturnsNoResult()
        {
            const string inputCode = @"
Public Function GetIt() As Object
End Function

Public Sub Baz()
    GetIt.Select
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        //See issue #5853 at https://github.com/rubberduck-vba/Rubberduck/issues/5853
        public void CallToMemberOfFunctionReturnValueInBody_NoResult()
        {
            const string returnTypeClassCode = @"
Public Sub Bar(ByVal arg As String)
End Sub
";
            const string moduleCode =
                @"
Public Function Foo() As Class1
    Set Foo = New Class1
    Foo.Bar ""SomeArgument""
End Function
";
            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", returnTypeClassCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        //See issue #5853 at https://github.com/rubberduck-vba/Rubberduck/issues/5853
        public void ExplicitCallToMemberOfFunctionReturnValueInBody_NoResult()
        {
            const string returnTypeClassCode = @"
Public Sub Bar(ByVal arg As String)
End Sub
";
            const string moduleCode =
                @"
Public Function Foo() As Class1
    Set Foo = New Class1
    Call Foo.Bar(""SomeArgument"")
End Function
";
            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", returnTypeClassCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void InterfaceMember_ReturnsResult()
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
            var modules = new (string, string, ComponentType)[]
            {
                ("IFoo", interfaceCode, ComponentType.ClassModule),
                ("Bar", implementationCode, ComponentType.ClassModule),
                ("TestModule", callSiteCode, ComponentType.StandardModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void MemberCallOnReturnValue_NoResult()
        {
            const string classCode = @"
Public Function Test() As Bar
End Function

Public Sub FooBar()
End Sub
";
            const string callSiteCode = @"
Public Sub Baz()
    Dim testObj As Bar
    Set testObj = new Bar
    testObj.Test.FooBar
End Sub
";
            var modules = new (string, string, ComponentType)[]
            {
                ("Bar", classCode, ComponentType.ClassModule),
                ("TestModule", callSiteCode, ComponentType.StandardModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void InterfaceMemberNotUsed_NoResult()
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
            var modules = new (string, string, ComponentType)[]
            {
                ("IFoo", interfaceCode, ComponentType.ClassModule),
                ("Bar", implementationCode, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("Unused Value")]
        public void InspectionName()
        {
            var inspection = InspectionUnderTest(null);

            Assert.AreEqual(nameof(FunctionReturnValueAlwaysDiscardedInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new FunctionReturnValueAlwaysDiscardedInspection(state);
        }
    }
}
