using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    
    [TestFixture]
    public class DeleteDeclarationsRefactoringActionTests
    {
        private static string threeConsecutiveNewLines = $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}";

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMultipleDeclarationTypes()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

Label1:    Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub

Public Property Get Test() As Long
    Test = 6
End Property

Public Sub DoNothing()
End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)

    Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub

Public Sub DoNothing()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "Test", "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [TestCase("var0", "Label1")]
        [TestCase("Label1", "var0")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMultipleDeclarationTypesLabelAndVariableAnyOrder(params string[] targets)
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

Label1:    Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)

var0 = arg

    Dim var2 As Variant
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [TestCase("Test", "localVar")]
        [TestCase("localVar", "Test")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFunctionAndLocalVariableAnyOrder(params string[] targets)
        {
            var inputCode =
@"
Private mTestVal As Long

Public Function Test() As Long
    Dim localVar As String
    localVar = ""asdf""
    Test = mTestVal
End Function

Public Sub DoNothing()
End Sub
";

            var expected =
@"
Private mTestVal As Long

Public Sub DoNothing()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, targets));
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void InvalidDeclarationType_throws()
        {
            var inputCode =
$@"
Option Explicit

Public Sub TestSub(arg1 As Long)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var refactoringAction = DeleteDeclarationsTestSupport.CreateDeleteDeclarationRefactoringAction(state, rewritingManager);
                Assert.Throws<InvalidDeclarationTypeException>(() => refactoringAction.Refactor(TestModel(state, "arg1"), rewritingManager.CheckOutCodePaneSession()));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void AddTargetsAfterModelInstantiation()
        {
            var inputCode =
$@"
Option Explicit

Private mVar0 As Long

Private mVar1 As Long

Private mVar2 As String

Private mVar31 As Integer

Private mVar32 As Boolean

Public Sub TestSub(arg1 As Long)
End Sub

Public Function Bizz() As Long
End Function
";
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var dVar0 = state.DeclarationFinder.MatchName("mVar0").First();

                var model = new DeleteDeclarationsModel(dVar0);

                var dVar1 = state.DeclarationFinder.MatchName("mVar1").First();
                var dVar2 = state.DeclarationFinder.MatchName("mVar2").First();

                model.AddDeclarationsToDelete(dVar1, dVar2);

                var multiple = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Where(d => d.IdentifierName.StartsWith("mVar3"));
                model.AddRangeOfDeclarationsToDelete(multiple);

                var funcBizz = state.DeclarationFinder.MatchName("Bizz").First();

                model.AddDeclarationsToDelete(funcBizz);
                var results = RefactoredCode(vbe, rewritingManager, state, model);
                var actualCode = results[MockVbeBuilder.TestModuleName];


                StringAssert.Contains("Public Sub TestSub(arg1 As Long)", actualCode);
                StringAssert.DoesNotContain("mVar", actualCode);
                StringAssert.DoesNotContain("Bizz", actualCode);
            }
        }

        //Every Enum must declare at least one member...Removing all members is equivalent to removing the entire Enum declaration
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllEnumMembers()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
    ThirdValue
End Enum

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "FirstValue", "SecondValue", "ThirdValue" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var model = new DeleteDeclarationsModel();
                model.AddRangeOfDeclarationsToDelete(testTargets);

                var results = RefactoredCode(vbe, rewritingManager, state, model);
                var actualCode = results[MockVbeBuilder.TestModuleName];

                StringAssert.Contains(expected, actualCode);
                StringAssert.AreEqualIgnoringCase(expected, actualCode);
            }
        }

        //Every UserDefinedType must declare at least one member...Removing all members is equivalent to removing the entire UDT declaration
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllUdtMembers()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As String
    ThirdValue As Boolean
End Type

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "FirstValue", "SecondValue", "ThirdValue" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var model = new DeleteDeclarationsModel();
                model.AddRangeOfDeclarationsToDelete(testTargets);

                var results = RefactoredCode(vbe, rewritingManager, state, model);
                var actualCode = results[MockVbeBuilder.TestModuleName];

                StringAssert.Contains(expected, actualCode);
                StringAssert.AreEqualIgnoringCase(expected, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllUdtMembersAndAllEnumMembers()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    EFirstValue
    ESecondValue
    EThirdValue
End Enum

Private Type TestType
    TFirstValue As Long
    TSecondValue As String
    TThirdValue As Boolean
End Type

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "EFirstValue", "ESecondValue", "EThirdValue", "TFirstValue", "TSecondValue", "TThirdValue" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var model = new DeleteDeclarationsModel();
                model.AddRangeOfDeclarationsToDelete(testTargets);

                var results = RefactoredCode(vbe, rewritingManager, state, model);
                var actualCode = results[MockVbeBuilder.TestModuleName];

                StringAssert.Contains(expected, actualCode);
                StringAssert.AreEqualIgnoringCase(expected, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMemberAndDeclarationUDT()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As String
    ThirdValue As Boolean
End Type

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "ThirdValue", "TestType" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var model = new DeleteDeclarationsModel();
                model.AddRangeOfDeclarationsToDelete(testTargets);

                var results = RefactoredCode(vbe, rewritingManager, state, model);
                var actualCode = results[MockVbeBuilder.TestModuleName];

                StringAssert.Contains(expected, actualCode);
                StringAssert.AreEqualIgnoringCase(expected, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMemberAndDeclarationEnum()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
    ThirdValue
End Enum

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "ThirdValue", "TestEnum" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var model = new DeleteDeclarationsModel();
                model.AddRangeOfDeclarationsToDelete(testTargets);

                var results = RefactoredCode(vbe, rewritingManager, state, model);
                var actualCode = results[MockVbeBuilder.TestModuleName];

                StringAssert.Contains(expected, actualCode);
                StringAssert.AreEqualIgnoringCase(expected, actualCode);
            }
        }

        private IDictionary<string, string> RefactoredCode(Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return RefactoredCode(vbe, modelBuilder);
        }

        private IDictionary<string, string> RefactoredCode(IVBE vbe, Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                return RefactoredCode(vbe, rewritingManager, state, modelBuilder(state));
            }
        }

        private IDictionary<string, string> RefactoredCode(IVBE vbe, IRewritingManager rewritingManager, IDeclarationFinderProvider finderProvider, DeleteDeclarationsModel model)
        {
            var refactoringAction = DeleteDeclarationsTestSupport.CreateDeleteDeclarationRefactoringAction(finderProvider, rewritingManager);

            var session = rewritingManager.CheckOutCodePaneSession();
            refactoringAction.Refactor(model, session);

            session.TryRewrite();

            return vbe.ActiveVBProject.VBComponents
                .ToDictionary(component => component.Name, component => component.CodeModule.Content());
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder)
        {
            var refactoredCode = RefactoredCode(
                modelBuilder,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static DeleteDeclarationsModel TestModel(RubberduckParserState state, params string[] identifiers)
        {
            var finder = state.DeclarationFinder;
            var targets = new List<Declaration>();
            foreach (var tgt in identifiers)
            {
                targets.Add(finder.MatchName(tgt).Single());
            }
            return new DeleteDeclarationsModel(targets);
        }
    }
}
