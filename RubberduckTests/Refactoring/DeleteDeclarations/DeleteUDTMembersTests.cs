using Antlr4.Runtime;
using Castle.Windsor;
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
    public class DeleteUDTMembersTests
    {
        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveUDTMemberDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type
";
            var modifiedDeclaration =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    SecondValue As Long
End Type
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"));
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.AreEqualIgnoringCase(modifiedDeclaration, actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        [TestCase("   ", "")]
        [TestCase("   ", "   ")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveUDTMemberDeclarationCommentBelow(string priorSpacing, string spacing)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Private Type TestType
{priorSpacing}FirstValue As Long  'This is the first value
{spacing}'This is the second value
    SecondValue As Long
End Type
";
            var modifiedDeclaration =
$@"
Option Explicit

Public mVar1 As Long

Private Type TestType
{spacing}'This is the second value
    SecondValue As Long
End Type
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"));
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.AreEqualIgnoringCase(modifiedDeclaration, actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        //Every UserDefinedType must have at least one member...Removing all members is equivalent to removing the entire UDT declaration
        //Passing all UDT members to the DeleteUDTMembersRefactoringAction results in an exception.  Passing them
        //in via the DeleteDeclarationsRefactoringAction removes the UserDefinedTypeDeclaration in its entirety.
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllUdtMembers_Throws()
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

Private Type TestType
'A comment preceding the FirstValue
    FirstValue As Long
'A comment preceding the SecondValue
    SecondValue As String
    ThirdValue As Boolean
End Type
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"), injectTODO);
            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;
            StringAssert.Contains($"'{injectedContent}A comment preceding the FirstValue", actualCode);
            StringAssert.Contains($"A comment preceding the SecondValue", actualCode);
        }

        private List<string> GetRetainedLines(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder, bool injectTODO = false)
            => GetRetainedCodeBlock(moduleCode, modelBuilder, injectTODO)
                .Trim()
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                .ToList();

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, bool injectTODO = false)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorUDTMembers,
                injectTODO,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorUDTMembers(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, bool injectTODOComment)
        {
            var model = new  DeleteUDTMembersModel(targets)
            {
                InsertValidationTODOForRetainedComments = injectTODOComment
            };

            var session = rewritingManager.CheckOutCodePaneSession();

            var refactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteUDTMembersRefactoringAction>();

            refactoringAction.Refactor(model, session);

            return session;
        }
    }
}
