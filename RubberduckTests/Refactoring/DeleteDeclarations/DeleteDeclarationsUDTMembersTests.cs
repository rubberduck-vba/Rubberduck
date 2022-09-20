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
    public class DeleteDeclarationsUDTMembersTests
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

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveUDTMemberDeclarationMultiple()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType 'TypeComment
'Comment About FirstValue or Type
    FirstValue As Long
    'FirstValue or SecondValue Comment
        SecondValue As Long
        'SecondValue or ThirdValue Comment
            ThirdValue As Long
            'ThirdValue Comment
                FourthValue As Long
                'FourthValue Comment
End Type
";
            var modifiedDeclaration =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType 'TypeComment
'Comment About FirstValue or Type
        'SecondValue or ThirdValue Comment
            ThirdValue As Long
            'ThirdValue Comment
                FourthValue As Long
                'FourthValue Comment
End Type
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue", "SecondValue"));
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.AreEqualIgnoringCase(modifiedDeclaration, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue"), (m) => m.InsertValidationTODOForRetainedComments = injectTODO);
            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;
            StringAssert.Contains($"{injectedContent}A comment preceding the FirstValue", actualCode);
            StringAssert.Contains($"A comment preceding the SecondValue", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeleteDeclarationsInMultipleUserDefinedTypes()
        {
            var inputCode =
@"
Option Explicit

Private Type TestType1
    FirstValue As Long
    SecondValue As String
End Type

Private Type TestType2
    FirstValue As Long
    SecondValue As String
End Type
";

            var expected =
@"
Option Explicit

Private Type TestType1
    SecondValue As String
End Type

Private Type TestType2
    SecondValue As String
End Type
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargetsUsingParentDeclaration(state, ("FirstValue", "TestType1"), ("FirstValue", "TestType2")));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
            StringAssert.Contains(expected, actualCode);
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, Action<IDeleteDeclarationsModel> modelFlagsAction = null)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorUDTMembers,
                modelFlagsAction ?? _support.DefaultModelFlagAction,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorUDTMembers(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, Action<IDeleteDeclarationsModel> modelFlagsAction)
        {
            var model = new DeleteUDTMembersModel(targets);
            modelFlagsAction(model);

            var session = rewritingManager.CheckOutCodePaneSession();

            var refactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteUDTMembersRefactoringAction>();

            refactoringAction.Refactor(model, session);

            return session;
        }
    }
}
