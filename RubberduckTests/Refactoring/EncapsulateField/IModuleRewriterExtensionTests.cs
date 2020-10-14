using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class IModuleRewriterExtensionTests
    {
        private static string threeConsecutiveNewLines = $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}";

        [Test]
        [Category("Rewriter")]
        public void RemoveFieldDeclarations()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Public mVar2 As Long

Private mVar3 As String, mVar4 As Long, mVar5 As String

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var content = TestRemoveBlocks(inputCode, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5");
            StringAssert.DoesNotContain(threeConsecutiveNewLines, content);
        }

        [Test]
        [Category("Rewriter")]
        public void RemoveFieldDeclarationsUsesColonStmtDelimiter()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long: Public mVar2 As Long

Private mVar3 As String, mVar4 As Long, mVar5 As String

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var content = TestRemoveBlocks(inputCode, "mVar1", "mVar3", "mVar4", "mVar5");
            StringAssert.DoesNotContain(threeConsecutiveNewLines, content);
            StringAssert.DoesNotContain(":", content);
        }

        [Test]
        [Category("Rewriter")]
        public void RemoveFieldDeclarations_LineContinuations()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Public mVar2 As Long

Private mVar3 As String _
        , mVar4 As Long _
                , mVar5 As String

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var content = TestRemoveBlocks(inputCode, "mVar1", "mVar2", "mVar3", "mVar5");
            StringAssert.DoesNotContain(threeConsecutiveNewLines, content);
        }

        [Test]
        [Category("Rewriter")]
        public void RemoveFieldDeclarations_RemovesAllBlankLines()
        {
            var inputCode =
@"
Option Explicit
Public mVar1 As Long





Public mVar2 As Long
Private mVar3 As String, mVar4 As Long, mVar5 As String





Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var content = TestRemoveBlocks(inputCode, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5");
            StringAssert.DoesNotContain(threeConsecutiveNewLines, content);
        }

        [Test]
        [Category("Rewriter")]
        public void RemoveMemberDeclarations()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub


Public Sub Test3()
End Sub

Public Sub Test4()
End Sub
";

            var content = TestRemoveBlocks(inputCode, "Test2", "Test3");
            StringAssert.DoesNotContain(threeConsecutiveNewLines, content);
        }

        private string TestRemoveBlocks(string inputCode, params string[] identifiers)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            (RubberduckParserState state, IRewritingManager rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var targets = new List<Declaration>();
                foreach (var id in identifiers)
                {
                    var target = state.DeclarationFinder
                        .MatchName(id).Single();
                    targets.Add(target);
                }

                var qmn = targets.First().QualifiedModuleName;
                var rewriteSession = rewriteManager.CheckOutCodePaneSession();
                var rewriter = rewriteSession.CheckOutModuleRewriter(qmn);

                rewriter.RemoveVariables(targets.OfType<VariableDeclaration>());
                rewriter.RemoveMembers(targets.OfType<ModuleBodyElementDeclaration>());

                return rewriter.GetText();
            }
        }
    }
}
