using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ParameterCanBeByValInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_NoResultForByValObjectInInterfaceImplementationProperty()
        {
            const string modelCode = @"
Option Explicit
Public Foo As Long
Public Bar As String
";

            const string interfaceCode = @"
Option Explicit

Public Property Get Model() As MyModel
End Property

Public Property Set Model(ByVal value As MyModel)
End Property

Public Property Get IsCancelled() As Boolean
End Property

Public Sub Show()
End Sub
";

            const string implementationCode = @"
Option Explicit
Private Type TView
    Model As MyModel
    IsCancelled As Boolean
End Type
Private this As TView
Implements IView

Private Property Get IView_IsCancelled() As Boolean
    IView_IsCancelled = this.IsCancelled
End Property

Private Property Set IView_Model(ByVal value As MyModel)
    Set this.Model = value
End Property

Private Property Get IView_Model() As MyModel
    Set IView_Model = this.Model
End Property

Private Sub IView_Show()
    Me.Show vbModal
End Sub
";
            var modules = new(string, string, ComponentType)[] 
            {
                ("IView", interfaceCode, ComponentType.ClassModule),
                ("MyModel", modelCode, ComponentType.ClassModule),
                ("MyForm", implementationCode, ComponentType.UserForm),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_NoResultForByValObjectInProperty()
        {
            const string inputCode =
                @"Public Property Set Foo(ByVal value As Object)
End Property";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_NoResultForByValObject()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 As Collection)
End Sub";
            var inspectionResults = InspectionResultsForStandardModule(inputCode);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestCase("Sub Foo(ByVal arg1 As Collection)\r\nEnd Sub", 0)]
        [TestCase("Sub Foo(arg1 As String)\r\nEnd Sub", 1)]
        [TestCase("Sub Foo(ByRef arg1 As String)\r\nEnd Sub", 1)]
        [TestCase("Sub Foo(arg1 As String, arg2 As Date)\r\nEnd Sub", 2)]
        [TestCase("Sub Foo(ByVal arg1 As String)\r\nEnd Sub", 0)]
        [TestCase("Sub Foo(arg1 As String, ByVal arg2 As Integer)\r\nEnd Sub", 1)]
        [TestCase("Sub Foo(ByRef arg1() As Variant)\r\nEnd Sub", 0)]
        [Category("Inspections")]
        public void ParameterCanBeByVal_NoResultForByValObject(string inputCode, int expectedCount)
        {
            var inspectionResults = InspectionResultsForStandardModule(inputCode);

            Assert.AreEqual(expectedCount, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedByRefAndAssigned()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void GivenByRefArrayParameter_ReturnsNoResult_Interface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a() As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a() As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void GivenBtRefArrayParameter_ReturnsNoResult_Event()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1() As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1() As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                 ("Class3", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedToByRefProc_NoAssignment()
        {
            const string inputCode =
                @"Sub DoSomething(foo As Integer)
    DoSomethingElse foo
End Sub

Sub DoSomethingElse(ByRef bar As Integer)
End Sub";
            Assert.IsFalse(InspectionResultsForStandardModule(inputCode).Any(result => result.Target.IdentifierName.Equals("foo")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedToByRefProc_WithAssignment()
        {
            const string inputCode =
                @"Sub DoSomething(foo As Integer)
    DoSomethingElse foo
End Sub

Sub DoSomethingElse(ByRef bar As Integer)
    bar = 42
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByRefProc_ExplicitlyByVal()
        {
            const string inputCode =
                @"Sub DoSomething(foo As Integer)
    DoSomethingElse (foo)
End Sub

Sub DoSomethingElse(ByRef bar As Integer)
    bar = 42
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByRefProc_PartOfExpression()
        {
            const string inputCode =
                @"Sub DoSomething(foo As Integer)
    DoSomethingElse foo + 2
End Sub

Sub DoSomethingElse(ByRef bar As Integer)
    bar = 42
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByValEvent()
        {
            const string inputCode =
                @" Public Event Bar(ByVal baz As Integer)

Sub DoSomething(foo As Integer)
    RaiseEvent Bar(foo)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedToByRefEvent()
        {
            const string inputCode =
                @" Public Event Bar(ByRef baz As Integer)

Sub DoSomething(foo As Integer)
    RaiseEvent Bar(foo)
End Sub";
            Assert.IsFalse(InspectionResultsForStandardModule(inputCode).Any(result => result.Target.IdentifierName.Equals("foo")));
        }


        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByRefEvent_ExplicilyByVal()
        {
            const string inputCode =
                @" Public Event Bar(ByRef baz As Integer)

Sub DoSomething(foo As Integer)
    RaiseEvent Bar(BYVAL foo)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count(result => result.Target.IdentifierName.Equals("foo")));
        }


        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByRefEvent_PartOfExpression()
        {
            const string inputCode =
                @" Public Event Bar(ByRef baz As Integer)

Sub DoSomething(foo As Integer)
    RaiseEvent Bar(foo + 2)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Where(result => result.Target.IdentifierName.Equals("foo")).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ParameterCanBeByVal
Sub Foo(arg1 As String)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParam()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleByValParam()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamAssignedTo_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
    a = 42
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamAssignedTo_InInterface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
    a = 42
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRefMethod_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
    DoSomething a
End Sub

Private Sub DoSomething(ByRef bar As Integer)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.IsFalse(InspectionResultsForModules(modules).Any(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRefMethod_InInterface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
    DoSomethingElse a
End Sub

Private Sub DoSomethingElse(ByRef bar As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.IsFalse(InspectionResultsForModules(modules).Any(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRefMethodExplicitlyByVal_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
    DoSomething (a)
End Sub

Private Sub DoSomething(ByRef bar As Integer)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRefMethodExplicitlyByVal_InInterface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
    DoSomethingElse (a)
End Sub

Private Sub DoSomethingElse(ByRef bar As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";
            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRefMethodPartOfExpression_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
    DoSomething a + 42
End Sub

Private Sub DoSomething(ByRef bar As Integer)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRefMethodPartOfExpression_InInterface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
    DoSomethingElse a + 42
End Sub

Private Sub DoSomethingElse(ByRef bar As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamByRefEvent_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Public Event Bar(ByRef foo As Integer)

Private Sub IClass1_DoSomething(ByRef a As Integer)
    RaiseEvent Bar(a)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.IsFalse(InspectionResultsForModules(modules).Any(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamByRefEvent_InInterface()
        {
            const string inputCode1 =
                @"Public Event Bar(ByRef foo As Integer)

Public Sub DoSomething(ByRef a As Integer)
    RaiseEvent Bar(a)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.IsFalse(InspectionResultsForModules(modules).Any(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamByRefEventExplicitlyByVal_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Public Event Bar(ByRef foo As Integer)

Private Sub IClass1_DoSomething(ByRef a As Integer)
    RaiseEvent Bar(ByVal a)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamByRefEventExplicitlyByVal_InInterface()
        {
            const string inputCode1 =
                @"Public Event Bar(ByRef foo As Integer)

Public Sub DoSomething(ByRef a As Integer)
    RaiseEvent Bar(ByVal a)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamByRefEventPartOfExpression_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Public Event Bar(ByRef foo As Integer)

Private Sub IClass1_DoSomething(ByRef a As Integer)
    RaiseEvent Bar(a + 15)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamByRefEventPartOfExpression_InInterface()
        {
            const string inputCode1 =
                @"Public Event Bar(ByRef foo As Integer)

Public Sub DoSomething(ByRef a As Integer)
    RaiseEvent Bar(a + 15)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("a")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_MultipleParams_OneCanBeByVal_InImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
           };

            Assert.AreEqual("a", InspectionResultsForModules(modules).Single().Target.IdentifierName);
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_MultipleParams_OneCanBeByVal_InInterface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
           {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual("a", InspectionResultsForModules(modules).Single().Target.IdentifierName);
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParam()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleByValParam()
        {
            const string inputCode1 =
                @"Public Event Foo(ByVal arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamAssignedTo()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamPassedToByRefProcedure()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    DoSomething arg1
End Sub

Private Sub DoSomething(ByRef bar As Integer)
End Sub
";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.IsFalse(InspectionResultsForModules(modules).Any(result => result.Target.IdentifierName.Equals("arg1")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamPassedToByRefProcedureExplicitlyByVal()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    DoSomething (arg1)
End Sub

Private Sub DoSomething(ByRef bar As Integer)
End Sub
";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("arg1")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamPassedToByRefProcedurePartOfExpression()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    DoSomething arg1 + 2
End Sub

Private Sub DoSomething(ByRef bar As Integer)
End Sub
";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("arg1")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamPassedToByRefRaiseEvent()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Public Event Bar(ByRef baz As Integer)

Private Sub abc_Foo(ByRef arg1 As Integer)
    RaiseEvent Bar(arg1)
End Sub
";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("arg1")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamPassedToByRefRaiseEventExplicitlyByVal()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Public Event Bar(ByRef baz As Integer)

Private Sub abc_Foo(ByRef arg1 As Integer)
    RaiseEvent Bar(ByVal arg1)
End Sub
";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("arg1")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamPassedToByRefRaiseEventPartOfExpression()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Public Event Bar(ByRef baz As Integer)

Private Sub abc_Foo(ByRef arg1 As Integer)
    RaiseEvent Bar(arg1 + 3)
End Sub
";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.IdentifierName.Equals("arg1")));
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EventMember_MultipleParams_OneCanBeByVal()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer, ByRef arg2 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer, ByRef arg2 As Integer)
    arg1 = 42
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual("arg2", InspectionResultsForModules(modules).Single().Target.IdentifierName);
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_EnumMemberParameterCanBeByVal()
        {
            const string inputCode = @"Option Explicit
Public Enum TestEnum
    Foo
    Bar
End Enum

Private Sub DoSomething(e As TestEnum)
    Debug.Print e
End Sub";

            Assert.AreEqual("e", InspectionResultsForStandardModule(inputCode).Single().Target.IdentifierName);
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_LibraryFunction_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Declare Function MyLibFunction Lib ""MyLib"" (arg1 As Integer) As Integer";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterCanBeByVal_LibraryProcedure_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Declare Sub MyLibProcedure Lib ""MyLib"" (arg1 As Integer)";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ParameterCanBeByValInspection(null);

            Assert.AreEqual(nameof(ParameterCanBeByValInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ParameterCanBeByValInspection(state);
        }
    }
}
