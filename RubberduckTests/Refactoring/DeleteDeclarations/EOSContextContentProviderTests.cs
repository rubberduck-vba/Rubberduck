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
    public class EOSContextContentProviderTests
    {
        private static string TestContent1 =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long) 'Is A DeclarationLogical Line Comment
    'First Pre-Annotation Comment Context
    'Second Pre-Annotation Comment Context
    '@Ignore VariableNotUsed, UseMeaningfulName
    'First Post-Annotation Comment Context
    'Second Post-Annotation Comment Context


    Dim X As Long
End Sub
";

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsDeclarationLogicalLineContext()
        {
            void thisTest(EOSContextContentProvider sut)
            {
                var goalContext = sut.DeclarationLogicalLineCommentContext;

                StringAssert.StartsWith(" 'Is A DeclarationLogical Line Comment", goalContext.GetText());
            }

            SetupAndInvokeTest(TestContent1, "DoSomethingElse", thisTest);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsDeclarationLogicalLineContextListDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public Const notUsed1 As Long = 100, _
    notUsed2 As Long = 200, _
        notUsed3 As Long = 300 _
            'These constants are not used


'This field is used
Public used As String
";

            void thisTest(EOSContextContentProvider sut)
            {
                var goalContext = sut.DeclarationLogicalLineCommentContext;
                var content = goalContext.GetText();
                StringAssert.Contains("'These constants are not used", content);
            }

            SetupAndInvokeTest(inputCode, "notUsed1", thisTest);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsDeclarationLogicalLineContextListDeclarationMultiLineComment()
        {
            var inputCode =
$@"
Option Explicit

    Private retained As Long

    'Comment above mVar1
Private mVar1 As Long 'This is a comment for mVar1 _
        'and so is this

        'Comment below mVar1
Public Sub Test()
End Sub";

            void thisTest(EOSContextContentProvider sut)
            {
                var goalContext = sut.DeclarationLogicalLineCommentContext;
                var content = goalContext.GetText();
                StringAssert.Contains("'This is a comment for mVar1", content);
                StringAssert.Contains("'and so is this", content);
            }

            SetupAndInvokeTest(inputCode, "mVar1", thisTest);
        }

        //[Test]
        //[Category("Refactorings")]
        //[Category(nameof(DeleteDeclarationsRefactoringAction))]
        //public void FindsPreAnnotationLineContexts()
        //{
        //    void thisTest(EOSContextContentProvider sut)
        //    {
        //        var goalContexts = sut.PreAnnotationCommentContexts;

        //        Assert.AreEqual(2, goalContexts.Count);
        //        StringAssert.Contains("'First Pre-Annotation Comment Context", goalContexts.First().GetText());
        //        var second = goalContexts.Skip(1).First();
        //        StringAssert.Contains("'Second Pre-Annotation Comment Context", second.GetText());
        //    }

        //    SetupAndInvokeTest(TestContent1, "DoSomethingElse", thisTest);
        //}

        //[Test]
        //[Category("Refactorings")]
        //[Category(nameof(DeleteDeclarationsRefactoringAction))]
        //public void FindsPostAnnotationLineContexts()
        //{
        //    void thisTest(EOSContextContentProvider sut)
        //    {
        //        var goalContexts = sut.PostAnnotationCommentContexts;

        //        Assert.AreEqual(2, goalContexts.Count);
        //        StringAssert.Contains("'First Post-Annotation Comment Context", goalContexts.First().GetText());
        //        var second = goalContexts.Skip(1).First();
        //        StringAssert.Contains("'Second Post-Annotation Comment Context", second.GetText());
        //    }

        //    SetupAndInvokeTest(TestContent1, "DoSomethingElse", thisTest);
        //}

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsSeparationToNextContext()
        {
            void thisTest(EOSContextContentProvider sut)
            {
                //StringAssert.Contains($"{Environment.NewLine}{ Environment.NewLine}", sut.EndingNewLines);
                StringAssert.Contains($"{Environment.NewLine}{ Environment.NewLine}", sut.Separation);
            }

            SetupAndInvokeTest(TestContent1, "DoSomethingElse", thisTest);
        }

        [TestCase("    ")]
        [TestCase( "")]
        [TestCase( "        ")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsIndentationToNextContext(string expectedIndentation)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long) 'Is A DeclarationLogical Line Comment
    'First Pre-Annotation Comment Context
    'Second Pre-Annotation Comment Context
    '@Ignore VariableNotUsed, UseMeaningfulName
    'First Post-Annotation Comment Context
    'Second Post-Annotation Comment Context


{expectedIndentation}Dim X As Long
End Sub
";
            void thisTest(EOSContextContentProvider sut)
            {
                //StringAssert.Contains(expectedIndentation, sut.NextContextIndentation);
                StringAssert.Contains(expectedIndentation, sut.Indentation);
            }

            SetupAndInvokeTest(inputCode, "DoSomethingElse", thisTest);
        }

        private static void SetupAndInvokeTest(string inputCode, string memberName, Action<EOSContextContentProvider> testSUT)
        {
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var xVar = state.DeclarationFinder.MatchName(memberName).First();
                var eos = GetEOS(xVar.Context);

                var rewriter = rewritingManager.CheckOutCodePaneSession().CheckOutModuleRewriter(xVar.QualifiedModuleName);

                 var eosEX = new EOSContextContentProvider(eos, rewriter);

                testSUT(eosEX);
            }
        }

        private static VBAParser.EndOfStatementContext GetEOS(ParserRuleContext targetContext)
        {
            switch (targetContext)
            {
                case VBAParser.ConstSubStmtContext _:
                case VBAParser.VariableSubStmtContext _:
                    targetContext.GetAncestor<VBAParser.ModuleDeclarationsElementContext>()
                        .TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var mdeEos);
                    if (mdeEos != null)
                    {
                        return mdeEos;
                    }
                    targetContext.GetAncestor<VBAParser.BlockStmtContext>()
                        .TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var blockStmtCtxt);
                    return blockStmtCtxt;
                default:
                    return targetContext.GetChild<VBAParser.EndOfStatementContext>();
            }
        }
    }
}
