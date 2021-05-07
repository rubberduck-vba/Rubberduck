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
    public class DeclarationsDeleter_EnumTests
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
        //Passing all enum members to the EnumMemberDeclarationsDeleter results in an exception.  Passing them
        //in via the DeleteDeclarationRefactoring substitutes the EnumDeclaration for the members.
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

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var testTargets = new List<Declaration>();
                foreach (var id in new string[] { "FirstValue", "SecondValue", "ThirdValue" })
                {
                    testTargets.Add(state.DeclarationFinder.MatchName(id).First());
                }

                var enumDeleter = new DeleteEnumMembersRefactoringAction(state, rewritingManager);
                var model = new DeleteEnumMembersModel(testTargets);

                Assert.Throws<InvalidOperationException>(() => enumDeleter.Refactor(model, rewritingManager.CheckOutCodePaneSession()));
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
                var session = rewritingManager.CheckOutCodePaneSession();

                var refactoringAction = new DeleteEnumMembersRefactoringAction(state, rewritingManager);

                var model = new DeleteEnumMembersModel(modelBuilder(state));

                refactoringAction.Refactor(model, session);

                session.TryRewrite();

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }
    }
}
