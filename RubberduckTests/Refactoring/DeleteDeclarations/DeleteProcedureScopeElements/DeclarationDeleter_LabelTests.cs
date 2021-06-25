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
    public class DeclarationDeleter_LabelTests
    {
        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [TestCase("var0 = arg")]
        [TestCase("Dim var1 As Long: var1 = arg")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveLabelWithFollowingExpression(string expression)
        {
            var inputCode =
$@"
Sub Foo(ByVal arg As Long)
    Dim var0 As Long
Label1:    {expression}

End Sub";

            var expected =
$@"
Sub Foo(ByVal arg As Long)
    Dim var0 As Long
    {expression}

End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [TestCase("Label1:    'Comment on Label1 line\r\n\r\n", "")]
        [TestCase("Label1:    Dim var0 As Long    'Comment on Label1 line\r\n\r\n", "    Dim var0 As Long    'Comment on Label1 line\r\n\r\n")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithSameLineContent(string testExpression, string expectedRewrite)
        {
            var inputCode =
$@"
Sub Foo(ByVal arg As Long)

{testExpression}End Sub";

            var expected =
$@"
Sub Foo(ByVal arg As Long)

{expectedRewrite}End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, bool injectTODO = false)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorProcedureScopeElements,
                injectTODO,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorProcedureScopeElements(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, bool injectTODOComment)
        {
            var model = new DeleteProcedureScopeElementsModel(targets)
            {
                InsertValidationTODOForRetainedComments = injectTODOComment
            };

            var session = rewritingManager.CheckOutCodePaneSession();

            var refactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteProcedureScopeElementsRefactoringAction>();

            refactoringAction.Refactor(model, session);

            return session;
        }
    }
}
