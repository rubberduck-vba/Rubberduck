using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class NonReturningFunctionInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningPropertyGet_ReturnsResult()
        {
            const string inputCode =
                @"Property Get Foo() As Boolean
End Property";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_MultipleFunctions()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function

Function Goo() As String
End Function";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_DoesNotReturnResult_Let()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Foo = True
End Function";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_DoesNotReturnResult_Set()
        {
            const string inputCode =
                @"Function Foo() As Collection
    Set Foo = new Collection
End Function";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore NonReturningFunction
Function Foo() As Boolean
End Function";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_MultipleSubs_SomeReturning()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Foo = True
End Function

Function Goo() As String
End Function";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_GivenParenthesizedByRefAssignment()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    ByRefAssign (Foo)
End Function

Public Sub ByRefAssign(ByRef a As Boolean)
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_GivenUseStrictlyInsideByRefAssignment()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    ByRefAssign Foo + 42
End Function

Public Sub ByRefAssign(ByRef a As Boolean)
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_NoResult_GivenByRefAssignment_WithMemberAccess()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    TestModule1.ByRefAssign False, Foo
End Function

Public Sub ByRefAssign(ByVal v As Boolean, ByRef a As Boolean)
    a = v
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_GivenUnassignedByRefAssignment_WithMemberAccess()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    TestModule1.ByRefAssign False, Foo
End Function

Public Sub ByRefAssign(ByVal v As Boolean, ByRef a As Boolean)
    'nope, not assigned
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_NoResult_GivenClassMemberCall()
        {
            const string code = @"
Public Function Foo() As Boolean
    With New Class1
        .ByRefAssign Foo
    End With
End Function
";
            const string classCode = @"
Public Sub ByRefAssign(ByRef b As Boolean)
End Sub
";
            var builder = new MockVbeBuilder();
            builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, classCode);
            var vbe = builder.Build();
            Assert.AreEqual(0, InspectionResults(vbe.Object).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_NoResult_GivenByRefAssignment_WithNamedArgument()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    ByRefAssign b:=Foo
End Function

Public Sub ByRefAssign(Optional ByVal a As Long, Optional ByRef b As Boolean)
    b = False
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //ref issue #5964
        public void NonReturningFunction_NoResult_AssignmenToUDTMembersInWithBlock()
        {
            const string inputCode = @"
Private Type tipo
    one As Long
    two As Long
End Type

Function assigner() As tipo
    With assigner
        .one = 1
        .two = 2
    End With
End Function
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //ref issue #5964
        public void NonReturningFunction_NoResult_AssignmenToUDTMembersInWithBlock_NestedWith_Inside()
        {
            const string inputCode = @"
Private Type tipo
    one As Long
    two As Long
End Type

Function assigner() As tipo
    Dim bar As tipo
    With bar
        .one = 3
        .two = 2
        With assigner
            .one = 1
            .two = 2
        End With
    End With
End Function
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //ref issue #5964
        public void NonReturningFunction_NoResult_AssignmenToUDTMembersInWithBlock_NestedWith_Start()
        {
            const string inputCode = @"
Private Type tipo
    one As Long
    two As Long
End Type

Function assigner() As tipo
    Dim bar As tipo
    With assigner
        .one = 3
        .two = 2
        With bar
            .one = 1
            .two = 2
        End With
    End With
End Function
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //ref issue #5964
        public void NonReturningFunction_NoResult_AssignmenToUDTMembersInWithBlock_NestedWith_End()
        {
            const string inputCode = @"
Private Type tipo
    one As Long
    two As Long
End Type

Function assigner() As tipo
    Dim bar As tipo
    With assigner
        With bar
            .one = 1
            .two = 2
        End With
        .one = 3
        .two = 2
    End With
End Function
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //ref issue #5964
        public void NonReturningFunction_OneResult_AssignmenToUDTMembersOfOtherVariableInNestedWith()
        {
            const string inputCode = @"
Private Type tipo
    one As Long
    two As Long
End Type

Function assigner() As tipo
    Dim bar As tipo
    With assigner
        With bar
            .one = 1
            .two = 2
        End With
    End With
End Function
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
                @"Function Foo() As Boolean
End Function";
            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Boolean
End Function";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new NonReturningFunctionInspection(null);

            Assert.AreEqual(nameof(NonReturningFunctionInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new NonReturningFunctionInspection(state);
        }
    }
}
