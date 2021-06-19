using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class DeleterEnumMembersTests
    {
        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveEnumMemberDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
End Enum

Public Sub Test1()
End Sub
";
            var modifiedDeclaration =
@"
Private Enum TestEnum
    SecondValue
End Enum
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        [TestCase("   ", "")]
        [TestCase("   ", "   ")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveEnumMemberDeclarationRetainsCommentBelow(string spacingToFirstElement, string spacing)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
{spacingToFirstElement}FirstValue  'This is the first value
{spacing}'This is the second value
    SecondValue
End Enum

Public Sub Test1()
End Sub
";
            var modifiedDeclaration =
$@"
Private Enum TestEnum
{spacing}'This is the second value
    SecondValue
End Enum
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        //Every Enum must have at least one member...Removing all members is equivalent to removing the entire Enum declaration
        //Passing all enum members to the DeleteEnumMembersRefactoringAction results in an exception.  Passing them
        //in via the DeleteDeclarationsRefactoringAction removes the EnumDeclaration in its entirety.
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllEnumMembers_Throws()
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

            Assert.Throws<InvalidOperationException>(() => GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue", "SecondValue", "ThirdValue")));
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsInjectTODOCommentFlag(bool injectTODO)
        {
            var inputCode =
@"
Option Explicit

Private Enum TestEnum
'A comment preceding the FirstValue
    FirstValue
'A comment preceding the SecondValue
    SecondValue
End Enum

";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"), injectTODO);
            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;
            StringAssert.Contains($"'{injectedContent}A comment preceding the FirstValue", actualCode);
            StringAssert.Contains($"A comment preceding the SecondValue", actualCode);
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, bool injectTODO = false)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorEnumMembers,
                injectTODO,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorEnumMembers(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, bool injectTODOComment)
        {
            var refactoringAction = new DeleteEnumMembersRefactoringAction(state, new DeclarationDeletionTargetFactory(state), rewritingManager);

            var model = new DeleteEnumMembersModel(targets)
            {
                InjectTODOForRetainedComments = injectTODOComment
            };

            var session = rewritingManager.CheckOutCodePaneSession();

            refactoringAction.Refactor(model, session);

            return session;
        }
    }
}
