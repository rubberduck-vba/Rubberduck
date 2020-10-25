using System;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractInterfaceTests : InteractiveRefactoringTestBase<IExtractInterfacePresenter, ExtractInterfaceModel>
    {
        private static readonly string errRaise = "Err.Raise 5";
        private static readonly string todoMsg = Rubberduck.Resources.Refactorings.Refactorings.ImplementInterface_TODO;

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProc()
        {
            var inputCode =
@"Public Sub Foo(ByVal a|rg1 As Integer, ByVal arg2 As String)
End Sub";
            var code = ToCodeAndSelection(inputCode);

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsAllMembers(model);

            var actualCode = RefactoredCode("Class", code.Selection, presenterAction, null, false, ("Class", code.Code, ComponentType.ClassModule));

            var classCode = actualCode["Class"];
            StringAssert.Contains("Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)", classCode);
            StringAssert.Contains(errRaise, classCode);
            StringAssert.Contains(todoMsg, classCode);

            var actualInterfaceCode = actualCode[actualCode.Keys.Single(componentName => !componentName.Equals("Class"))];
            StringAssert.Contains("Option Explicit", actualInterfaceCode);
            StringAssert.Contains("'@Interface", actualInterfaceCode);
            StringAssert.Contains("Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)", actualInterfaceCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_InvalidDeclarationTypeException_Throws()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsAllMembers(model);

            var refactoredCode = RefactoredCode(
                "Module",
                DeclarationType.Module,
                presenterAction,
                typeof(InvalidDeclarationTypeException),
                ("Module", inputCode, ComponentType.StandardModule));

            Assert.AreEqual(inputCode, refactoredCode["Module"]);
            Assert.AreEqual(1, refactoredCode.Count);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NoDeclarationForSelectionException()
        {
            const string inputCode =
@"

|
Public Sub Foo()
End Sub";
            var code = ToCodeAndSelection(inputCode);

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsAllMembers(model);

            var refactoredCode = RefactoredCode(
                "Module",
                code.Selection,
                presenterAction,
                typeof(NoDeclarationForSelectionException),
                false,
                ("Module", code.Code, ComponentType.StandardModule));

            Assert.AreEqual(code.Code, refactoredCode["Module"]);
            Assert.AreEqual(1, refactoredCode.Count);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_InvalidRefactoringModelException_Throws()
        {
            var inputCode =
@"Private Sub Fo|o()
End Sub";
            var code = ToCodeAndSelection(inputCode);

            var refactoredCode = RefactoredCode(
                "Class",
                code.Selection,
                model => null,
                typeof(InvalidRefactoringModelException),
                false,
                ("Class", code.Code, ComponentType.ClassModule));

            Assert.AreEqual(code.Code, refactoredCode["Class"]);
            Assert.AreEqual(1, refactoredCode.Count);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_IgnoresField()
        {
            var model = RetrieveModel("Public Fi|zz As Boolean".ToCodeString());
            Assert.AreEqual(0, model.Members.Count);
        }

        [TestCase("Attribute VB_Exposed = True", ClassInstancing.Public)]
        [TestCase("", ClassInstancing.Private)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_DefaultsToPublicInterface(string attribute, ClassInstancing expected )
        {
            var inputCode =
$@"{attribute}

Public Sub F|oo
End Sub";

            var model = RetrieveModel(inputCode.ToCodeString());
            Assert.AreEqual(expected, model.InterfaceInstancing);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NullPresenter_NoChanges()
        {
            var inputCode =
@"Private Sub Fo|o()
End Sub";
            var code = ToCodeAndSelection(inputCode);

            var vbe = MockVbeBuilder.BuildFromSingleModule(code.Code, ComponentType.ClassModule, out var component, code.Selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), code.Selection);

                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IExtractInterfacePresenter, ExtractInterfaceModel>(It.IsAny<ExtractInterfaceModel>())).Returns(value: null);

                var selectionService = MockedSelectionService();

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                Assert.AreEqual(1, vbe.Object.ActiveVBProject.VBComponents.Count());
                Assert.AreEqual(code.Code, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PassTargetIn()
        {
            var inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsFirstMember(model);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var classCode = actualCode["Class"];
            StringAssert.Contains("Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)", classCode);
            StringAssert.Contains(errRaise, classCode);
            StringAssert.Contains(todoMsg, classCode);

            var actualInterfaceCode = actualCode[actualCode.Keys.Single(componentName => !componentName.Equals("Class"))];
            StringAssert.Contains("Option Explicit", actualInterfaceCode);
            StringAssert.Contains("'@Interface", actualInterfaceCode);
            StringAssert.Contains("Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)", actualInterfaceCode);
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_Subroutine(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Long

Public Sub Fizz(ByVal arg1 As Integer, ByVal arg2 As String)
    mFizz = arg1 * CLng(arg2)
End Sub
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Sub IClass_Fizz(ByVal arg1 As Integer, ByVal arg2 As String)", sourceModuleCode);
            StringAssert.Contains($"mFizz = arg1 * CLng(arg2){Environment.NewLine}", sourceModuleCode);
            StringAssert.DoesNotContain($"mFizz = arg1 * CLng(arg2){Environment.NewLine}{Environment.NewLine}", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains($"IClass_Fizz arg1, arg2{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain($"IClass_Fizz arg1, arg2{Environment.NewLine}{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains($"Fizz arg1, arg2{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain($"Fizz arg1, arg2{Environment.NewLine}{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Sub Fizz(ByVal arg1 As Integer, ByVal arg2 As String)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PropertyLet(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Long

Public Property Let Fizz(value As Long)
    mFizz = value
End Property
";
            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Let IClass_Fizz(ByVal value As Long)", sourceModuleCode);
            StringAssert.Contains($"mFizz = value{Environment.NewLine}", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains($"IClass_Fizz = value{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains($"Fizz = value{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Property Let Fizz(value As Long)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ParameterizedPropertyLetWithParameters(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Long

Public Property Let Fizz(arg1 As Integer, arg2 As Integer, value As Long)
    mFizz = value
End Property
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Let IClass_Fizz(arg1 As Integer, arg2 As Integer, ByVal value As Long)", sourceModuleCode);
            StringAssert.Contains("mFizz = value", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains($"IClass_Fizz(arg1, arg2) = value{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains($"Fizz(arg1, arg2) = value{Environment.NewLine}", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Property Let Fizz(arg1 As Integer, arg2 As Integer, value As Long)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_Function(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Long

Public Function Fizz(ByVal arg1 As Integer, ByVal arg2 As String) As Long
    mFizz = arg1 * CLng(arg2)
    Fizz = mFizz
End Function
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Function IClass_Fizz(ByVal arg1 As Integer, ByVal arg2 As String)", sourceModuleCode);
            StringAssert.Contains($"mFizz = arg1 * CLng(arg2){Environment.NewLine}", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Fizz = IClass_Fizz(arg1, arg2)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("IClass_Fizz = Fizz(arg1, arg2)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Function Fizz(ByVal arg1 As Integer, ByVal arg2 As String)", sourceModuleCode);
                    StringAssert.Contains("IClass_Fizz = mFizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PropertySet(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Variant

Public Property Set Fizz(value As Variant)
    Set mFizz = value
End Property
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Set IClass_Fizz(ByVal value As Variant)", sourceModuleCode);
            StringAssert.Contains("Set mFizz = value", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Set IClass_Fizz = value", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("Set Fizz = value", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Property Set Fizz(value As Variant)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PropertyGet(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Long

Public Property Get Fizz() As Long
    Fizz = mFizz
End Property
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Get IClass_Fizz() As Long", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Fizz = IClass_Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("IClass_Fizz = Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Property Get Fizz() As Long", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PropertyGetObject(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As CTest

Private Sub Class_Initialize()
    Set mFizz = new CTest
End Sub

Public Property Get Fizz() As CTest
    Set Fizz = mFizz
End Property
";

            var cTestCode =
@"
Option Explicit

Public Sub Fizz()
End Sub
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule), ("TestClass", cTestCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Get IClass_Fizz() As CTest", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Set Fizz = IClass_Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("Set IClass_Fizz = Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_FunctionObject(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As CTest

Private Sub Class_Initialize()
    Set mFizz = new CTest
End Sub

Public Function Fizz() As CTest
    Set Fizz = mFizz
End Function
";

            var cTestCode =
@"
Option Explicit

Public Sub Fizz()
End Sub
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule), ("TestClass", cTestCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Function IClass_Fizz() As CTest", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Set Fizz = IClass_Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("Set IClass_Fizz = Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PropertyGetVariant(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Variant

Public Property Get Fizz() As Variant
    If IsObject(mFizz) Then
        Set Fizz = mFizz
    Else
        Fizz = mFizz
    End If
End Property
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Get IClass_Fizz() As Variant", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Fizz = IClass_Fizz", sourceModuleCode);
                    StringAssert.Contains("Set IClass_Fizz = mFizz", sourceModuleCode);
                    StringAssert.Contains("    IClass_Fizz = mFizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("IClass_Fizz = Fizz", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.NoInterfaceImplementation)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers)]
        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ParameterizedPropertyGet(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Private mFizz As Long

Public Property Get Fizz(arg1 As Long, arg2 As Long) As Long
    Fizz = mFizz
End Property
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("Private Property Get IClass_Fizz(arg1 As Long, arg2 As Long) As Long", sourceModuleCode);
            switch (extractOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                    StringAssert.Contains("Fizz = IClass_Fizz(arg1, arg2)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    StringAssert.Contains("IClass_Fizz = Fizz(arg1, arg2)", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    StringAssert.DoesNotContain("Public Property Get Fizz(arg1 As Long, arg2 As Long) As Long", sourceModuleCode);
                    StringAssert.DoesNotContain(errRaise, sourceModuleCode);
                    break;
                case ExtractInterfaceImplementationOption.NoInterfaceImplementation:
                    StringAssert.Contains(errRaise, sourceModuleCode);
                    break;
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_RenamesReferences(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Public Function AddABunch(ByVal arg1 As Long) As Long
    AddABunch = AddOne(arg1) + AddTwo(arg1) + AddThree(arg1)
End Function

Public Function AddOne(ByVal arg1 As Long) As Long
    AddOne = arg1 + 1
End Function

Public Function AddTwo(ByVal arg1 As Long) As Long
    AddTwo = arg1 + 2
End Function

Public Function AddThree(ByVal arg1 As Long) As Long
    AddThree = arg1 + 3
End Function
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsAllMembers(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.DoesNotContain(errRaise, sourceModuleCode);
            StringAssert.Contains("IClass_AddABunch = IClass_AddOne(arg1) + IClass_AddTwo(arg1) + IClass_AddThree(arg1)", sourceModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_RemovesBlankLines()
        {
            const string inputCode =
@"
Public Function AddOne(ByVal arg1 As Long) As Long
    AddOne = arg1 + 1
End Function

Public Function AddTwo(ByVal arg1 As Long) As Long
    AddTwo = arg1 + 2
End Function

Public Function AddThree(ByVal arg1 As Long) As Long
    AddThree = arg1 + 3
End Function

Public Function AddFour(ByVal arg1 As Long) As Long
    AddThree = arg1 + 4
End Function

Public Function AddFive(ByVal arg1 As Long) As Long
    AddThree = arg1 + 5
End Function
";

            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                foreach (var member in model.Members)
                {
                    member.IsSelected = !(member.Member.IdentifierName.EndsWith("One") || member.Member.IdentifierName.EndsWith("Five"));
                }
                model.ImplementationOption = ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface;
                return model;
            };

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.DoesNotContain("Function AddTwo(", sourceModuleCode);
            StringAssert.DoesNotContain("Function AddThree(", sourceModuleCode);
            StringAssert.DoesNotContain("Function AddFour(", sourceModuleCode);
            StringAssert.Contains("Function IClass_AddTwo(", sourceModuleCode);
            StringAssert.Contains("Function IClass_AddThree(", sourceModuleCode);
            StringAssert.Contains("Function IClass_AddFour(", sourceModuleCode);
            StringAssert.Contains($"End Function{Environment.NewLine}{Environment.NewLine}Public Function", sourceModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_RemovesBlankLines2()
        {
            const string inputCode =
@"
Option Explicit

Private mName As String
Private mTestVariant As Variant

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(value As String)
    mName = value
End Property

Public Property Get TestVariant() As Variant
    If IsObject(mTestVariant) Then
        Set TestVariant = mTestVariant
    Else
         TestVariant = mTestVariant
    End If
End Property

Public Property Let TestVariant(value As Variant)
    mTestVariant = value
End Property

Public Property Set TestVariant(value As Variant)
    Set mTestVariant = value
End Property

Private Sub Class_Initialize()
    Name = ""Bill""
End Sub


Public Sub MySpecialSub()

End Sub

Public Sub MyMultiArgSub(arg1 As Long, arg2 As String)
    Name = arg2
End Sub
";
            var expected =
@"
Option Explicit

Implements IClass

Private mName As String
Private mTestVariant As Variant

Private Sub Class_Initialize()
    IClass_Name = ""Bill""
End Sub

Private Property Get IClass_Name()";

            var extractOption = ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface;

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsAllMembers(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.StartsWith(expected, sourceModuleCode);
        }

        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_EmptyBodyHasTODO(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Public Function AddABunch(ByVal arg1 As Long) As Long
    AddABunch = AddOne(arg1) + AddTwo(arg1) + AddThree(arg1)
End Function

Public Function AddOne(ByVal arg1 As Long) As Long
    AddOne = arg1 + 1
End Function

Public Function AddTwo(ByVal arg1 As Long) As Long
End Function

Public Function AddThree(ByVal arg1 As Long) As Long
    AddThree = arg1 + 3
End Function
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsAllMembers(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.Contains("IClass_AddABunch = IClass_AddOne(arg1) + IClass_AddTwo(arg1) + IClass_AddThree(arg1)", sourceModuleCode);

            var emptyBody =
$@"Private Function IClass_AddTwo(ByVal arg1 As Long) As Long{Environment.NewLine}    Err.Raise 5";
            StringAssert.Contains(emptyBody, sourceModuleCode);

            if (extractOption == ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)
            {
                var forwardBody =
    $@"Public Function AddTwo(ByVal arg1 As Long) As Long{Environment.NewLine}    AddTwo = IClass_AddTwo";
                StringAssert.Contains(forwardBody, sourceModuleCode);
            }
        }

        [TestCase(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)]
        [TestCase(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_RenamesReferencesOfNonInterfaceMembers(ExtractInterfaceImplementationOption extractOption)
        {
            const string inputCode =
@"
Public Function AddABunch(ByVal arg1 As Long) As Long
    AddABunch = AddOne(arg1) + AddTwo(arg1) + AddThree(arg1)
End Function

Public Function AddOne(ByVal arg1 As Long) As Long
    AddOne = arg1 + 1
End Function

Public Function AddTwo(ByVal arg1 As Long) As Long
    AddTwo = arg1 + 2
End Function

Public Function AddThree(ByVal arg1 As Long) As Long
    AddThree = arg1 + 3
End Function

Public Function SomeOtherFunction(ByVal arg1 As Long) As Long
    SomeOtherFunction = AddOne(arg1) + AddTwo(arg1) + AddThree(arg1)
End Function
";

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model)
                => UserSelectsAllMembers(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];
            StringAssert.DoesNotContain(errRaise, sourceModuleCode);
            StringAssert.Contains("IClass_AddABunch = IClass_AddOne(arg1) + IClass_AddTwo(arg1) + IClass_AddThree(arg1)", sourceModuleCode);
            StringAssert.Contains("SomeOtherFunction = IClass_AddOne(arg1) + IClass_AddTwo(arg1) + IClass_AddThree(arg1)", sourceModuleCode);
        }

        [Test]
        [Category("Extract Interface")]
        [TestCase("Label:")]
        [TestCase("Const Bar = 42")]
        [TestCase("Dim bar As Long")]
        [TestCase("Const Bar = 42: Dim baz As Long")]
        [TestCase("Const Bar = 42\nDim baz As Long")]
        [TestCase("Label: Const Bar = 42: Dim baz As Long")]
        [TestCase("Label:\nConst Bar = 42\nDim baz As Long")]
        public void ExtractInterfaceRefactoring_VariousNonExecutableContent(string statement)
        {
            var inputCode =
$@"Sub Foo()
    {statement}
End Sub";
            var extractOption = ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface;

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsFirstMember(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"];

            StringAssert.Contains("IClass_Foo()", sourceModuleCode);
            StringAssert.Contains($"IClass_Foo(){Environment.NewLine}", sourceModuleCode);
            var statementParts = statement.Split(new char[] { '\n' });
            foreach (var part in statementParts)
            {
                StringAssert.Contains($"{part}", sourceModuleCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_IndenterApplied()
        {
            var inputCode =
$@"
Public Function DivideBy(ByVal arg1 As Long, ByVal arg2 As Long) As Single
On Error Goto ErrorExit:
Dim result As Single
result = arg1 / arg2
DivideBy = result
Exit Function
ErrorExit:
DivideBy = 0
End Function";

            var expectedCode =
$@"Private Function IClass_DivideBy(ByVal arg1 As Long, ByVal arg2 As Long) As Single
    On Error Goto ErrorExit:
    Dim result As Single
    result = arg1 / arg2
    IClass_DivideBy = result
    Exit Function
    ErrorExit:
    IClass_DivideBy = 0
End Function";

            var extractOption = ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface;

            ExtractInterfaceModel presenterAction(ExtractInterfaceModel model) 
                => UserSelectsAllMembers(model, extractOption);

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            var sourceModuleCode = actualCode["Class"].Substring(actualCode["Class"].IndexOf("\r\n")).Trim();
            Assert.AreEqual(expectedCode, sourceModuleCode);
        }

        private static ExtractInterfaceModel UserSelectsFirstMember(ExtractInterfaceModel model, ExtractInterfaceImplementationOption implementationOption = ExtractInterfaceImplementationOption.NoInterfaceImplementation)
        {
            model.Members.ElementAt(0).IsSelected = true;
            model.ImplementationOption = implementationOption;
            return model;
        }

        private static ExtractInterfaceModel UserSelectsAllMembers(ExtractInterfaceModel model, ExtractInterfaceImplementationOption implementationOption = ExtractInterfaceImplementationOption.NoInterfaceImplementation)
        {
            foreach (var member in model.Members)
            {
                member.IsSelected = true;
            }
            model.ImplementationOption = implementationOption;
            return model;
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IExtractInterfacePresenter, ExtractInterfaceModel> userInteraction,
            ISelectionService selectionService)
        {
            var addImplementationsBaseRefactoring = new AddInterfaceImplementationsRefactoringAction(rewritingManager, CreateCodeBuilder());
            var addComponentService = TestAddComponentService(state?.ProjectsProvider);
            var baseRefactoring = new ExtractInterfaceRefactoringAction(addImplementationsBaseRefactoring, state, state, rewritingManager, state?.ProjectsProvider, addComponentService);
            var conflictFinderFactory = new TestConflictFinderFactory() as IExtractInterfaceConflictFinderFactory;
            return new ExtractInterfaceRefactoring(baseRefactoring, state, userInteraction, selectionService, conflictFinderFactory, CreateCodeBuilder());
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }

        private class TestConflictFinderFactory : IExtractInterfaceConflictFinderFactory
        {
            public IExtractInterfaceConflictFinder Create(IDeclarationFinderProvider declarationFinderProvider, string projectId)
            {
                return new ExtractInterfaceConflictFinder(declarationFinderProvider, projectId);
            }
        }

        private static ExtractInterfaceModel RetrieveModel(TestCodeString codeString)
            => RetrieveModel(codeString.Code, codeString.CaretPosition.ToOneBased());

        private static ExtractInterfaceModel RetrieveModel(string code, Selection selection)
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule(code, ComponentType.ClassModule, out _, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var target = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ClassModule)
                    .OfType<ClassModuleDeclaration>()
                    .First();

                return new ExtractInterfaceModel(state, target, CreateCodeBuilder());
            }
        }

        private static (string Code, Selection Selection) ToCodeAndSelection(string input)
        {
            var codeString = input.ToCodeString();
            return (codeString.Code, codeString.CaretPosition.ToOneBased());
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