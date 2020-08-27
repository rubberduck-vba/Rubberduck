using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.CodeBlockInsert;
using RubberduckTests.Mocks;
using System.Linq;

namespace RubberduckTests.Refactoring.CodeBlockInsert
{
    [TestFixture]
    public class CodeBlockInsertRefactoringActionTests : RefactoringActionTestBase<CodeBlockInsertModel>
    {
        private static string _declarationSectionMisplaced = "DeclarationSection content in wrong location";
        private static string _codeSectionMisplaced = "CodeSection content in wrong location";
        private static string _commentMisplaced = "PostContentMessage content in wrong location";

        private static string _declarationContent = "'This is the DeclarationSection";
        private static string _codeContent = "'This is the CodeSection";
        private static string _commentContent = "'This is a comment";

        [Test]
        [Category("Refactorings")]
        [Category(nameof(CodeBlockInsertRefactoringAction))]
        public void CorrectOrder()
        {
            string inputCode =
$@"
Option Explicit

";
            var results = ExecuteTest(inputCode, (_declarationContent, _codeContent, _commentContent));
            var declarationsIndex = results.IndexOf(_declarationContent);
            var codeIndex = results.IndexOf(_codeContent);
            var commentIndex = results.IndexOf(_commentContent);
            Assert.Greater(codeIndex, declarationsIndex, _codeSectionMisplaced);
            Assert.Greater(commentIndex, codeIndex, _commentMisplaced);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(CodeBlockInsertRefactoringAction))]
        public void CorrectOrder_NoMembers()
        {
            string inputCode =
$@"
Option Explicit

Private mTest As Long

";
            var results = ExecuteTest(inputCode, (_declarationContent, _codeContent, _commentContent));
            var existingDeclarationIndex = results.IndexOf("Private mTest As Long");
            var declarationsIndex = results.IndexOf(_declarationContent);
            var codeIndex = results.IndexOf(_codeContent);
            var commentIndex = results.IndexOf(_commentContent);
            Assert.Greater(declarationsIndex, existingDeclarationIndex, _declarationSectionMisplaced);
            Assert.Greater(codeIndex, declarationsIndex, _codeSectionMisplaced);
            Assert.Greater(commentIndex, codeIndex, _commentMisplaced);
        }


        [Test]
        [Category("Refactorings")]
        [Category(nameof(CodeBlockInsertRefactoringAction))]
        public void CorrectOrder_ExistingMember()
        {
            string inputCode =
$@"
Option Explicit

Private mTest As Long

Public Sub Test()
End Sub
";
            var results = ExecuteTest(inputCode, (_declarationContent, _codeContent, _commentContent));
            var existingDeclarationIndex = results.IndexOf("Private mTest As Long");
            var existingMemberIndex = results.IndexOf("Sub Test()");
            var declarationsIndex = results.IndexOf(_declarationContent);
            var codeIndex = results.IndexOf(_codeContent);
            var commentIndex = results.IndexOf(_commentContent);
            Assert.Greater(declarationsIndex, existingDeclarationIndex, _declarationSectionMisplaced);
            Assert.Greater(codeIndex, declarationsIndex, _codeSectionMisplaced);
            Assert.Greater(existingMemberIndex, codeIndex, _codeSectionMisplaced);
            Assert.Greater(existingMemberIndex, commentIndex, _commentMisplaced);
        }

        private string ExecuteTest(string inputCode, params (string, string, string)[] content)
        {
            return RefactoredCode(inputCode, state => TestModel(state, content));
        }

        private CodeBlockInsertModel TestModel(RubberduckParserState state, params (string declarationSection, string codeSection, string comment)[] blocks)
        {
            var members = state.DeclarationFinder.UserDeclarations(DeclarationType.Member);

            var module = state.DeclarationFinder.MatchName(MockVbeBuilder.TestModuleName).Single();
            var model = new CodeBlockInsertModel()
            {
                QualifiedModuleName = module.QualifiedModuleName,
                CodeSectionStartIndex = members
                    .OrderBy(m => m?.Context.Start.TokenIndex)
                    .FirstOrDefault()?.Context.Start.TokenIndex,
            };

            foreach ((string declaration, string code, string comment) in blocks)
            {
                model.AddContentBlock(NewContentType.DeclarationBlock, declaration);
                model.AddContentBlock(NewContentType.CodeSectionBlock, code);
                model.AddContentBlock(NewContentType.PostContentMessage, comment);
            }
            return model;
        }

        protected override IRefactoringAction<CodeBlockInsertModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new CodeBlockInsertRefactoringAction(state, rewritingManager, new CodeBuilder());
        }
    }
}