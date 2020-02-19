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

Private Property Let IClass_Buzz(value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Set IClass_Buzz(value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property
";

            const string expectedInterfaceCode =
                @"Option Explicit

'@Interface

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(ByRef b As Variant) As Variant
End Function

Public Property Get Buzz() As Variant
End Property

Public Property Let Buzz(ByRef value As Variant)
End Property

Public Property Set Buzz(ByRef value As Variant)
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

Public Function Fizz(ByRef b As Variant) As Variant
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

        private void ExecuteTest(string inputCode, string expectedClassCode, string expectedInterfaceCode, Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment)
        {
            var refactoredCode = RefactoredCode(
                state => TestModel(state, modelAdjustment),
                ("Class", inputCode, ComponentType.ClassModule));

            Assert.AreEqual(expectedClassCode, refactoredCode["Class"]);
            Assert.AreEqual(expectedInterfaceCode, refactoredCode["IClass"]);
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
            var model = new ExtractInterfaceModel(state, targetClass);
            return modelAdjustment(model);
        }

        protected override IRefactoringAction<ExtractInterfaceModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var addInterfaceImplementationsAction = new AddInterfaceImplementationsRefactoringAction(rewritingManager);
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