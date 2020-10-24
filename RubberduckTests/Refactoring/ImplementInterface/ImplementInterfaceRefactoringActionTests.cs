using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Settings;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ImplementInterfaceRefactoringActionTests : RefactoringActionTestBase<ImplementInterfaceModel>
    {
        private string _errorRaiseStmt = "Err.Raise 5";
        private string _todoStmt = "'TODO implement interface member";
        private string ErrRaiseAndComment => $"{_errorRaiseStmt}  {_todoStmt}";

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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
End Property

Private Property Let Interface1_a(ByVal RHS As String)
    {ErrRaiseAndComment}
End Property

Private Property Get Interface1_b() As String
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
End Sub

Private Function Interface1_Bar(ByVal a As Integer) As Boolean
    {ErrRaiseAndComment}
End Function

Private Property Get Interface1_Buz(ByVal a As Boolean) As Integer
    {ErrRaiseAndComment}
End Property

Private Property Let Interface1_Buz(ByVal a As Boolean, ByVal value As Integer)
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
End Sub

Private Function Interface1_Fizz(b As Variant) As Variant
    {ErrRaiseAndComment}
End Function

Private Property Get Interface1_Buzz() As Variant
    {ErrRaiseAndComment}
End Property

Private Property Let Interface1_Buzz(ByVal value As Variant)
    {ErrRaiseAndComment}
End Property

Private Property Set Interface1_Buzz(ByVal value As Variant)
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
End Property

Private Property Let Interface1_Foo(ByVal {_rhsIdentifier} As Long)
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
End Property

Private Property Set Interface1_Foo(ByVal {_rhsIdentifier} As Object)
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
End Property

Private Property Let Interface1_Foo(ByVal {_rhsIdentifier} As Variant)
    {ErrRaiseAndComment}
End Property

Private Property Set Interface1_Foo(ByVal {_rhsIdentifier} As Variant)
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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
    {ErrRaiseAndComment}
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

            //Remove Indenter formatting effects from refactoring results evaluation
            var expected = expectedClassCode.Trim().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var refactored = refactoredCode["Class1"].Trim().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            Assert.AreEqual(expected.Count(), refactored.Count());
            for (var idx = 0; idx < expected.Count(); idx++)
            {
                if (expected[idx].Contains(_errorRaiseStmt))
                {
                    StringAssert.Contains(_errorRaiseStmt, refactored[idx]);
                    StringAssert.Contains(_todoStmt, refactored[idx]);
                    continue;
                }
                Assert.AreEqual(expected[idx], refactored[idx]);
            }
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
            var addInterfaceImplementationsAction = new AddInterfaceImplementationsRefactoringAction(rewritingManager, CreateCodeBuilder());
            return new ImplementInterfaceRefactoringAction(addInterfaceImplementationsAction, rewritingManager);
        }

        private static ICodeBuilder CreateCodeBuilder()
            => new CodeBuilder(new Indenter(null, CreateIndenterSettings));

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }
    }
}