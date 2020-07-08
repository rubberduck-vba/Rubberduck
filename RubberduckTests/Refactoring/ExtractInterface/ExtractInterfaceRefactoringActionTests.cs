using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractInterfaceRefactoringActionTests : RefactoringActionTestBase<ExtractInterfaceModel>
    {
        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProc()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProcAndFuncAndPropGetSetLet()
        {
            //Input
            const string inputCode = @"
Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            //Expectation
            const string expectedClassCode = @"Implements IClass


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function IClass_Fizz(b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function

Private Property Get IClass_Buzz() As Variant
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IClass_Buzz(ByVal value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Set IClass_Buzz(ByVal value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b As Variant) As Variant
End Function

Public Property Get Buzz() As Variant
End Property

Public Property Let Buzz(ByVal value As Variant)
End Property

Public Property Set Buzz(ByVal value As Variant)
End Property
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProcAndFunc_IgnoreProperties()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function IClass_Fizz(b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b As Variant) As Variant
End Function
";
            var modelAdjustment = SelectFilteredMembers(member => !member.FullMemberSignature.Contains("Property"));
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, modelAdjustment);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_BelowLastImplementStatement()
        {
            //Input
            const string inputCode =
                @"

Option Explicit 

Implements Interface1


Implements Interface2



Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Expectation
            const string expectedClassCode =
                @"

Option Explicit 

Implements Interface1


Implements Interface2
Implements IClass



Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_BelowLastOptionStatement()
        {
            //Input
            const string inputCode =
                @"

Option Explicit 



Option Base 1





Private bar As Variant


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Expectation
            const string expectedClassCode =
                @"

Option Explicit 



Option Base 1

Implements IClass





Private bar As Variant


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_AtTopOfModule()
        {
            //Input
            const string inputCode =
                @"











Private bar As Variant


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Expectation
            const string expectedClassCode =
                @"Implements IClass













Private bar As Variant


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PutsInterfaceInFolderOfClassItIsExtractedFrom()
        {
            //Input
            const string inputCode =
                @"'@Folder(""MyFolder.MySubFolder"")

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Expectation
            const string expectedClassCode =
                @"Implements IClass

'@Folder(""MyFolder.MySubFolder"")

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Folder(""MyFolder.MySubFolder"")
'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_ImplicitByRefParameter()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(arg As Variant)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(arg As Variant)
End Sub

Private Sub IClass_Foo(arg As Variant)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(arg As Variant)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_ExplicitByRefParameter()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByRef arg As Variant)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(ByRef arg As Variant)
End Sub

Private Sub IClass_Foo(ByRef arg As Variant)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByRef arg As Variant)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_ByValParameter()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg As Variant)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(ByVal arg As Variant)
End Sub

Private Sub IClass_Foo(ByVal arg As Variant)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg As Variant)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_OptionalParameter_WoDefault()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(Optional arg As Variant)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(Optional arg As Variant)
End Sub

Private Sub IClass_Foo(Optional arg As Variant)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(Optional arg As Variant)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_OptionalParameter_WithDefault()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(Optional arg As Variant = 42)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(Optional arg As Variant = 42)
End Sub

Private Sub IClass_Foo(Optional arg As Variant = 42)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(Optional arg As Variant = 42)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_ParamArray()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(arg1 As Long, ParamArray args() As Variant)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(arg1 As Long, ParamArray args() As Variant)
End Sub

Private Sub IClass_Foo(arg1 As Long, ParamArray args() As Variant)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(arg1 As Long, ParamArray args() As Variant)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_MakesMissingAsTypesExplicit()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(arg1)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(arg1)
End Sub

Private Sub IClass_Foo(arg1 As Variant)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(arg1 As Variant)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_Array()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(arg1() As Long)
End Sub";
            //Expectation
            const string expectedClassCode =
                @"Implements IClass

Public Sub Foo(arg1() As Long)
End Sub

Private Sub IClass_Foo(arg1() As Long)
    Err.Raise 5 'TODO implement interface member
End Sub
";
            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(arg1() As Long)
End Sub
";
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, SelectAllMembers);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ExtractInterfaceRefactoring_PublicInterfaceInstancingCreatesExposedInterface()
        {

            //Input
            const string inputCode =
                @"'@Folder(""MyFolder.MySubFolder"")

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Expectation
            const string expectedClassCode =
                @"Implements IClass

'@Folder(""MyFolder.MySubFolder"")

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""IClass""
Attribute VB_Exposed = True
Option Explicit

'@Folder(""MyFolder.MySubFolder"")
'@Exposed
'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment = model =>
            {
                var modifiedModel = SelectAllMembers(model);
                modifiedModel.InterfaceInstancing = ClassInstancing.Public;
                return modifiedModel;
            };
            ExecuteTest(inputCode, expectedClassCode, expectedInterfaceCode, modelAdjustment);
        }

        private void ExecuteTest(string inputCode, string expectedClassCode, string expectedInterfaceCode, Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment)
        {
            var refactoredCode = RefactoredCode(
                state => TestModel(state, modelAdjustment),
                ("Class", inputCode, ComponentType.ClassModule));

            Assert.AreEqual(expectedClassCode.Trim(), refactoredCode["Class"].Trim());
            Assert.AreEqual(expectedInterfaceCode.Trim(), refactoredCode["IClass"].Trim());
        }

        private static ExtractInterfaceModel SelectAllMembers(ExtractInterfaceModel model)
        {
            foreach (var interfaceMember in model.Members)
            {
                interfaceMember.IsSelected = true;
            }

            return model;
        }

        private static Func<ExtractInterfaceModel, ExtractInterfaceModel> SelectFilteredMembers(Func<InterfaceMember, bool> filter)
        {
            return model => SelectFilteredMembers(model, filter);
        }

        private static ExtractInterfaceModel SelectFilteredMembers(ExtractInterfaceModel model, Func<InterfaceMember, bool> filter)
        {
            foreach (var interfaceMember in model.Members.Where(filter))
            {
                interfaceMember.IsSelected = true;
            }

            return model;
        }

        private static ExtractInterfaceModel TestModel(IDeclarationFinderProvider state, Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment)
        {
            var finder = state.DeclarationFinder;
            var targetClass = finder.UserDeclarations(DeclarationType.ClassModule)
                .OfType<ClassModuleDeclaration>()
                .Single(module => module.IdentifierName == "Class");
            var model = new ExtractInterfaceModel(state, targetClass, new CodeBuilder());
            return modelAdjustment(model);
        }

        protected override IRefactoringAction<ExtractInterfaceModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var addInterfaceImplementationsAction = new AddInterfaceImplementationsRefactoringAction(rewritingManager, new CodeBuilder());
            var addComponentService = TestAddComponentService(state?.ProjectsProvider);
            return new ExtractInterfaceRefactoringAction(addInterfaceImplementationsAction, state, state, rewritingManager, state?.ProjectsProvider, addComponentService);
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }
    }
}