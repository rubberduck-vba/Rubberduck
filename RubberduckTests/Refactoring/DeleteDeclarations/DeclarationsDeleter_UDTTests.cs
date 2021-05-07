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
    public class DeclarationsDeleter_UDTTests
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
        //Passing all UDT members to the UdtMemberDeclarationsDeleter results in an exception.  Passing them
        //in via the DeleteDeclarationRefactoring substitutes the EnumDeclaration for the members.
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

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "FirstValue", "SecondValue", "ThirdValue" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var udtDeleter = new DeleteEnumMembersRefactoringAction(state, rewritingManager);
                var model = new DeleteEnumMembersModel(testTargets);

                Assert.Throws<InvalidOperationException>(() => udtDeleter.Refactor(model, rewritingManager.CheckOutCodePaneSession()));
            }
        }

        private List<string> GetRetainedLines(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder)
            => GetRetainedCodeBlock(moduleCode, modelBuilder)
                .Trim()
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                .ToList();

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder)
        {
            var refactoredCode = ModifiedCode(
                modelBuilder,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private IDictionary<string, string> ModifiedCode(Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return ModifiedCode(vbe, modelBuilder);
        }

        private static IDictionary<string, string> ModifiedCode(IVBE vbe, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var refactoringAction = new DeleteUDTMembersRefactoringAction(state, rewritingManager);

                var session = rewritingManager.CheckOutCodePaneSession();

                var model = new DeleteUDTMembersModel(modelBuilder(state));

                refactoringAction.Refactor(model, session);

                session.TryRewrite();

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }
    }
}
