using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Settings;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractInterfaceRefactoringActionTests : RefactoringActionTestBase<ExtractInterfaceModel>
    {
        private static readonly string errRaise = "Err.Raise 5";
        private static readonly string todoMsg = Rubberduck.Resources.Refactorings.Refactorings.ImplementInterface_TODO;

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProc()
        {
            var inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains("Implements IClass", actualCode);
            StringAssert.Contains("IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)", actualCode);
            StringAssert.Contains(errRaise, actualCode);
            StringAssert.Contains(todoMsg, actualCode);

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains("'@Interface", interfaceModule);
            StringAssert.Contains("Foo(ByVal arg1 As Integer, ByVal arg2 As String)", interfaceModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProcAndFuncAndPropGetSetLet()
        {
            var inputCode =
@"
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

            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains("Implements IClass", actualCode);
            StringAssert.Contains("Get IClass_Buzz() As Variant", actualCode);
            StringAssert.Contains("Let IClass_Buzz(ByVal value As Variant)", actualCode);
            StringAssert.Contains("Set IClass_Buzz(ByVal value As Variant)", actualCode);
            StringAssert.Contains("IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)", actualCode);
            StringAssert.Contains("Function IClass_Fizz(b As Variant)", actualCode);

            Assert.AreEqual(5, CountOfStringByLine(actualCode, errRaise));
            Assert.AreEqual(5, CountOfStringByLine(actualCode, todoMsg));

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains("'@Interface", interfaceModule);
            StringAssert.Contains("Foo(ByVal arg1 As Integer, ByVal arg2 As String)", interfaceModule);
            StringAssert.Contains("Function Fizz(b As Variant)", interfaceModule);
            StringAssert.Contains("Get Buzz() As Variant", interfaceModule);
            StringAssert.Contains("Let Buzz(ByVal value As Variant)", interfaceModule);
            StringAssert.Contains("Set Buzz(ByVal value As Variant)", interfaceModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProcAndFunc_IgnoreProperties()
        {
            var inputCode =
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
            var modelAdjustment = SelectFilteredMembers(member => !member.FullMemberSignature.Contains("Property"));
            var refactoredCode = RefactoredCode(inputCode, modelAdjustment);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains("Implements IClass", actualCode);
            StringAssert.Contains("IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)", actualCode);
            StringAssert.Contains("Function IClass_Fizz(b As Variant)", actualCode);

            Assert.AreEqual(2, CountOfStringByLine(actualCode, errRaise));
            Assert.AreEqual(2, CountOfStringByLine(actualCode, todoMsg));

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains("'@Interface", interfaceModule);
            StringAssert.Contains("Foo(ByVal arg1 As Integer, ByVal arg2 As String)", interfaceModule);
            StringAssert.Contains("Function Fizz(b As Variant)", interfaceModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_BelowLastImplementStatemens()
        {
            var inputCode =
@"
Option Explicit 

Implements Interface1


Implements Interface2
Public Sub Foo()
End Sub";

            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            Assert.IsTrue(StringsFoundInThisOrderByLine(actualCode, "Implements Interface2", "Implements IClass", "Public Sub Foo()"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_BelowLastOptionStatement()
        {
            var inputCode =
@"
Option Explicit 

Option Base 1

Private bar As Variant
";

            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            Assert.IsTrue(StringsFoundInThisOrderByLine(actualCode, "Option Base 1", "Implements IClass", "Private bar As Variant"));
        }

        [TestCase(1)]
        [TestCase(10)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_AtTopOfModule(int newLineCount)
        {
            var newLines = string.Concat(Enumerable.Repeat(Environment.NewLine, newLineCount));
            var inputCode =
$@"{newLines}
Private bar As Variant
";
            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var lines = refactoredCode["Class"].Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            Assert.IsTrue(lines[0].Equals("Implements IClass"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PutsInterfaceInFolderOfClassItIsExtractedFrom()
        {
            var inputCode =
@"'@Folder(""MyFolder.MySubFolder"")

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains(@"'@Folder(""MyFolder.MySubFolder"")", actualCode);

            var interfaceModule = refactoredCode["IClass"];
            Assert.IsTrue(StringsFoundInThisOrderByLine(interfaceModule, @"'@Folder(""MyFolder.MySubFolder"")", "'@Interface"));
        }

        [TestCase("ByVal ")]
        [TestCase("ByRef ")]
        [TestCase("")]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ParameterMechanism(string parameterMechanism)
        {
            var inputCode =
$@"Public Sub Foo({parameterMechanism}arg As Variant)
End Sub";
            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains($"IClass_Foo({parameterMechanism}arg As Variant)", actualCode);

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains($"Foo({parameterMechanism}arg As Variant)", interfaceModule);
        }

        [TestCase("Optional arg As Variant")]
        [TestCase("Optional arg As Variant = 42")]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_OptionalParameter(string optionalParameter)
        {
            var inputCode =
$@"Public Sub Foo({optionalParameter})
End Sub";
            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains($"IClass_Foo({optionalParameter})", actualCode);

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains($"Foo({optionalParameter})", interfaceModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ParamArray()
        {
            var inputCode =
@"Public Sub Foo(arg1 As Long, ParamArray args() As Variant)
End Sub";
            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains("IClass_Foo(arg1 As Long, ParamArray args() As Variant)", actualCode);

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains("Foo(arg1 As Long, ParamArray args() As Variant)", interfaceModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_MakesMissingAsTypesExplicit()
        {
            var inputCode =
@"Public Sub Foo(arg1)
End Sub";
            var refactoredCode = RefactoredCode(inputCode, SelectAllMembers);

            var actualCode = refactoredCode["Class"];
            StringAssert.Contains("Sub IClass_Foo(arg1 As Variant)", actualCode);

            var interfaceModule = refactoredCode["IClass"];
            StringAssert.Contains("Sub Foo(arg1 As Variant)", interfaceModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_Array()
        {
            var subDeclaration = "Foo(arg1() As Long)";
            var targetModuleName = "Class";

            var inputCode =
$@"Public Sub {subDeclaration}
End Sub";

            var refactoredCode = RefactoredCode(
                state => TestModel(state, SelectAllMembers, targetModuleName),
                (targetModuleName, inputCode, ComponentType.ClassModule));

            var sourceModuleCode = refactoredCode[targetModuleName];
            StringAssert.Contains(subDeclaration, sourceModuleCode);
            StringAssert.Contains($"Implements I{targetModuleName}", sourceModuleCode);
            StringAssert.Contains($"I{targetModuleName}_{subDeclaration}", sourceModuleCode);
            StringAssert.Contains(errRaise, sourceModuleCode);

            var interfaceClassCode = refactoredCode[$"I{targetModuleName}"];
            StringAssert.Contains(subDeclaration, interfaceClassCode);
            StringAssert.Contains("'@Interface", interfaceClassCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PublicInterfaceInstancingCreatesExposedInterface()
        {

            var inputCode =
@"'@Folder(""MyFolder.MySubFolder"")

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var expectedInterfaceCode =
@"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""IClass""
Attribute VB_Exposed = True
'@Folder(""MyFolder.MySubFolder"")
'@Exposed
'@Interface

Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment = model =>
            {
                var modifiedModel = SelectAllMembers(model);
                modifiedModel.InterfaceInstancing = ClassInstancing.Public;
                modifiedModel.ImplementationOption = ExtractInterfaceImplementationOption.NoInterfaceImplementation;
                return modifiedModel;
            };

            var interfaceModule = RefactoredCode(inputCode, modelAdjustment)["IClass"];

            StringAssert.AreEqualIgnoringCase(expectedInterfaceCode.Trim(), interfaceModule.Trim());
        }

        [TestCase("  mTest = 5\n")]
        [TestCase(null)]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ReplaceMemberWithInterface(string content)
        {
            var testModuleName = "Class";
            var testModuleCode =
$@"Option Explicit

Private mTest As Long
    
Public Sub TestSub()
{content}End Sub
    
Public Sub TestSub1()
{content}End Sub
    
Public Sub TestSub2()
{content}End Sub
";
            var otherModuleName = "OtherModule";
            var otherModuleCode =
$@"Option Explicit

Public Sub ReferencingSub()
    Dim testClass As {testModuleName}
    Set testClass = {Tokens.New} {testModuleName}
    testClass.TestSub1
End Sub";

            var vbe = TestVbe((testModuleName, testModuleCode, ComponentType.ClassModule),
                (otherModuleName, otherModuleCode, ComponentType.StandardModule));

            Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment = model =>
            {
                var modifiedModel = SelectAllMembers(model);
                modifiedModel.ImplementationOption = ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface;
                return modifiedModel;
            };

            var refactoredCode = RefactoredCode(
               state => TestModel(state, modelAdjustment, testModuleName),
               (testModuleName, testModuleCode, ComponentType.ClassModule),
               (otherModuleName, otherModuleCode, ComponentType.StandardModule));

            var testModuleRefactoredCode = refactoredCode[testModuleName];
            StringAssert.Contains("Public Sub TestSub1()", testModuleRefactoredCode);
            StringAssert.DoesNotContain("Public Sub TestSub()", testModuleRefactoredCode);
            StringAssert.DoesNotContain("Public Sub TestSub2()", testModuleRefactoredCode);

            StringAssert.Contains("Private Sub IClass_TestSub()", testModuleRefactoredCode);
            StringAssert.Contains("Private Sub IClass_TestSub1()", testModuleRefactoredCode);
            StringAssert.Contains("Private Sub IClass_TestSub2()", testModuleRefactoredCode);

            if (content != null)
            {
                var body = content.Trim();
                Assert.AreEqual(3, CountOfStringByLine(testModuleRefactoredCode, body));
            }
            else
            {
                Assert.AreEqual(2, CountOfStringByLine(testModuleRefactoredCode, errRaise));
                Assert.AreEqual(3, CountOfStringByLine(testModuleRefactoredCode, todoMsg));
            }
        }

        private IDictionary<string, string> RefactoredCode(string inputCode, Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment, string targetModuleName = "Class")
        {
            return RefactoredCode(
               state => TestModel(state, modelAdjustment),
               (targetModuleName, inputCode, ComponentType.ClassModule));
        }

        private static ExtractInterfaceModel SelectAllMembers(ExtractInterfaceModel model)
        {
            foreach (var interfaceMember in model.Members)
            {
                interfaceMember.IsSelected = true;
            }

            model.ImplementationOption = ExtractInterfaceImplementationOption.NoInterfaceImplementation;
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
            model.ImplementationOption = ExtractInterfaceImplementationOption.NoInterfaceImplementation;
            return model;
        }

        private static ExtractInterfaceModel TestModel(IDeclarationFinderProvider state, Func<ExtractInterfaceModel, ExtractInterfaceModel> modelAdjustment, string testModuleName = "Class")
        {
            var finder = state.DeclarationFinder;
            var targetClass = finder.UserDeclarations(DeclarationType.ClassModule)
                .OfType<ClassModuleDeclaration>()
                .Single(module => module.IdentifierName == testModuleName);
            var model = new ExtractInterfaceModel(state, targetClass, CreateCodeBuilder());
            return modelAdjustment(model);
        }

        protected override IRefactoringAction<ExtractInterfaceModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return CreateRefactoringAction(state, rewritingManager);
        }

        private static IRefactoringAction<ExtractInterfaceModel> CreateRefactoringAction(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var addInterfaceImplementationsAction = new AddInterfaceImplementationsRefactoringAction(rewritingManager, CreateCodeBuilder());
            var addComponentService = TestAddComponentService(state?.ProjectsProvider);
            return new ExtractInterfaceRefactoringAction(addInterfaceImplementationsAction, state, state, rewritingManager, state?.ProjectsProvider, addComponentService);
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }

        private static int CountOfStringByLine(string actualCode, string target) 
            => actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                .Where(s => s.Contains(target)).Count();

        private static bool StringsFoundInThisOrderByLine(string actualCode, params string[] sequence)
        {
            var lines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();
            var idxSequence = 0;
            for (var idx = 0; idx < lines.Count && idxSequence < sequence.Count(); idx++)
            {
                if (lines[idx].Contains(sequence[idxSequence]))
                {
                    idxSequence++;
                }
            }
            return idxSequence > sequence.Count() - 1;
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