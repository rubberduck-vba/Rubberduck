using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ImplementInterfaceRefactoringActionTests : RefactoringActionTestBase<ImplementInterfaceModel>
    {
        private string _todoImplementMessage = "Err.Raise 5 'TODO implement interface member";

        private static string _rhsIdentifier = Rubberduck.Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam;

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Procedure()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo()
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo()
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Procedure_ClassHasOtherProcedure()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo()
End Sub";

            const string classCode =
                @"Implements Interface1

Public Sub Bar()
End Sub";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Public Sub Bar()
End Sub

Private Sub Interface1_Foo()
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PartiallyImplementedInterface()
        {
            //Input
            const string interfaceCode =
                @"Public Property Get a() As String
End Property
Public Property Let a(RHS As String)
End Property
Public Property Get b() As String
End Property
Public Property Let b(RHS As String)
End Property";

            const string classCode =
                @"Implements Interface1

Private Property Let Interface1_b(RHS As String)
End Property";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Let Interface1_b(RHS As String)
End Property

Private Property Get Interface1_a() As String
    {_todoImplementMessage}
End Property

Private Property Let Interface1_a(ByVal RHS As String)
    {_todoImplementMessage}
End Property

Private Property Get Interface1_b() As String
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Procedure_WithParams()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(ByVal a As Integer, ByRef b, c, d As Long)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(ByVal a As Integer, ByRef b As Variant, c As Variant, d As Long)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Function()
        {
            //Input
            const string interfaceCode =
                @"Public Function Foo() As Integer
End Function";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Function Interface1_Foo() As Integer
    {_todoImplementMessage}
End Function
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Function_WithImplicitType()
        {
            //Input
            const string interfaceCode =
                @"Public Function Foo()
End Function";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Function Interface1_Foo() As Variant
    {_todoImplementMessage}
End Function
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Function_WithParam()
        {
            //Input
            const string interfaceCode =
                @"Public Function Foo(a)
End Function";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Function Interface1_Foo(a As Variant) As Variant
    {_todoImplementMessage}
End Function
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertyGet()
        {
            //Input
            const string interfaceCode =
                @"Public Property Get Foo() As Integer
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Get Interface1_Foo() As Integer
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertyGet_WithImplicitType()
        {
            //Input
            const string interfaceCode =
                @"Public Property Get Foo()
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Get Interface1_Foo() As Variant
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertyGet_WithParam()
        {
            //Input
            const string interfaceCode =
                @"Public Property Get Foo(a)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Get Interface1_Foo(a As Variant) As Variant
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertyLet()
        {
            //Input
            const string interfaceCode =
                @"Public Property Let Foo(ByRef value As Long)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Let Interface1_Foo(ByVal value As Long)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertyLet_WithParam()
        {
            //Input
            const string interfaceCode =
                @"Public Property Let Foo(a)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Let Interface1_Foo(ByVal a As Variant)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertySet()
        {
            //Input
            const string interfaceCode =
                @"Public Property Set Foo(ByRef value As Variant)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Set Interface1_Foo(ByVal value As Variant)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertySet_WithParam()
        {
            //Input
            const string interfaceCode =
                @"Public Property Set Foo(a)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                 $@"Implements Interface1

Private Property Set Interface1_Foo(ByVal a As Variant)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PropertySet_AllTypes()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo()
End Sub

Public Function Bar(ByVal a As Integer) As Boolean
End Function

Public Property Get Buz(ByVal a As Boolean) As Integer
End Property

Public Property Let Buz(ByVal a As Boolean, ByRef value As Integer)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo()
    {_todoImplementMessage}
End Sub

Private Function Interface1_Bar(ByVal a As Integer) As Boolean
    {_todoImplementMessage}
End Function

Private Property Get Interface1_Buz(ByVal a As Boolean) As Integer
    {_todoImplementMessage}
End Property

Private Property Let Interface1_Buz(ByVal a As Boolean, ByVal value As Integer)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void CreatesMethodStubForAllProcedureKinds()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b)
End Function

Public Property Get Buzz() As Variant
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    {_todoImplementMessage}
End Sub

Private Function Interface1_Fizz(b As Variant) As Variant
    {_todoImplementMessage}
End Function

Private Property Get Interface1_Buzz() As Variant
    {_todoImplementMessage}
End Property

Private Property Let Interface1_Buzz(ByVal value As Variant)
    {_todoImplementMessage}
End Property

Private Property Set Interface1_Buzz(ByVal value As Variant)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [TestCase(@"Public Foo As Long")]
        [TestCase(@"Dim Foo As Long")]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PublicIntrinsic(string interfaceCode)
        {
            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Get Interface1_Foo() As Long
    {_todoImplementMessage}
End Property

Private Property Let Interface1_Foo(ByVal {_rhsIdentifier} As Long)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [TestCase(@"Public Foo As Object")]
        [TestCase(@"Dim Foo As Object")]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PublicObject(string interfaceCode)
        {
            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Get Interface1_Foo() As Object
    {_todoImplementMessage}
End Property

Private Property Set Interface1_Foo(ByVal {_rhsIdentifier} As Object)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [TestCase(@"Public Foo As Variant")]
        [TestCase(@"Public Foo")]
        [TestCase(@"Dim Foo As Variant")]
        [TestCase(@"Dim Foo")]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_PublicVariant(string interfaceCode)
        {
            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Property Get Interface1_Foo() As Variant
    {_todoImplementMessage}
End Property

Private Property Let Interface1_Foo(ByVal {_rhsIdentifier} As Variant)
    {_todoImplementMessage}
End Property

Private Property Set Interface1_Foo(ByVal {_rhsIdentifier} As Variant)
    {_todoImplementMessage}
End Property
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_ImplicitByRefParameter()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(arg As Variant)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(arg As Variant)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_ExplicitByRefParameter()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(ByRef arg As Variant)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(ByRef arg As Variant)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_ByValParameter()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(ByVal arg As Variant)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(ByVal arg As Variant)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_OptionalParameter_WoDefault()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(Optional arg As Variant)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(Optional arg As Variant)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_OptionalParameter_WithDefault()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(Optional arg As Variant = 42)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(Optional arg As Variant = 42)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_ParamArray()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(arg1 As Long, ParamArray args() As Variant)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(arg1 As Long, ParamArray args() As Variant)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_MakesMissingAsTypesExplicit()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(arg1)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                $@"Implements Interface1

Private Sub Interface1_Foo(arg1 As Variant)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_Array()
        {
            //Input
            const string interfaceCode =
                @"Public Sub Foo(arg1() As Long)
End Sub";

            const string classCode =
                @"Implements Interface1";

            //Expectation
            string expectedCode =
                 $@"Implements Interface1

Private Sub Interface1_Foo(arg1() As Long)
    {_todoImplementMessage}
End Sub
";
            ExecuteTest(classCode, interfaceCode, expectedCode);
        }

        private void ExecuteTest(string classCode, string interfaceCode, string expectedClassCode)
        {
            var refactoredCode = RefactoredCode(
                TestModel, 
                ("Class1", classCode,ComponentType.ClassModule),
                ("Interface1", interfaceCode, ComponentType.ClassModule));

            Assert.AreEqual(expectedClassCode.Trim(), refactoredCode["Class1"].Trim());
        }

        private static ImplementInterfaceModel TestModel(RubberduckParserState state)
        {
            var finder = state.DeclarationFinder;
            var targetInterface = finder.UserDeclarations(DeclarationType.ClassModule)
                .OfType<ClassModuleDeclaration>()
                .Single(module => module.IdentifierName == "Interface1");
            var targetClass = finder.UserDeclarations(DeclarationType.ClassModule)
                .OfType<ClassModuleDeclaration>()
                .Single(module => module.IdentifierName == "Class1");
            return new ImplementInterfaceModel(targetInterface, targetClass);
        }

        protected override IRefactoringAction<ImplementInterfaceModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var addInterfaceImplementationsAction = new AddInterfaceImplementationsRefactoringAction(rewritingManager, new CodeBuilder());
            return new ImplementInterfaceRefactoringAction(addInterfaceImplementationsAction, rewritingManager);
        }
    }
}