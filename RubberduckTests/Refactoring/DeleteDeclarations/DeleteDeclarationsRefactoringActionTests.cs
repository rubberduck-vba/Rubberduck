using System;
using System.Collections.Generic;
using System.Linq;
using Castle.Windsor;
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

        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Test", "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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
            Assert.Throws<InvalidDeclarationTypeException>(() => GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "arg1")));
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DecalarationsDeleteModel_AddDeclarationsToDelete()
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

                var deleteDeclarationsResolver = new DeleteDeclarationsTestsResolver(state, rewritingManager);
                var deleteDeclarationRefactoringAction = deleteDeclarationsResolver.Resolve<DeleteDeclarationsRefactoringAction>();

                var session = rewritingManager.CheckOutCodePaneSession();
                deleteDeclarationRefactoringAction.Refactor(model, session);

                session.TryRewrite();

                var results = vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());

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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue", "SecondValue", "ThirdValue"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue", "SecondValue", "ThirdValue"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, 
                state => _support.TestTargets(state, "EFirstValue", "ESecondValue", "EThirdValue", "TFirstValue", "TSecondValue", "TThirdValue"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "ThirdValue", "TestType"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "ThirdValue", "TestEnum"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, bool injectTODO = false)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorAllElementTypes,
                injectTODO,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorAllElementTypes(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, bool injectTODOComment)
        {
            var model = new DeleteDeclarationsModel(targets)
            {
                InsertValidationTODOForRetainedComments = injectTODOComment
            };

            var session = rewritingManager.CheckOutCodePaneSession();

            var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteDeclarationsRefactoringAction>();
            
            deleteDeclarationRefactoringAction.Refactor(model, session);

            return session;
        }
    }
}
