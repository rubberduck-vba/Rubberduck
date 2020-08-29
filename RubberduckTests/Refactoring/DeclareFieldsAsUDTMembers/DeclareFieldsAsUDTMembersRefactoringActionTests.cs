using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeclareFieldsAsUDTMembers;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveFieldsToUDT
{
    [TestFixture]
    public class DeclareFieldsAsUDTMembersRefactoringActionTests : RefactoringActionTestBase<DeclareFieldsAsUDTMembersModel>
    {
        [TestCase(4)]
        [TestCase(2)]
        [Category("Refactorings")]
        [Category(nameof(DeclareFieldsAsUDTMembersRefactoringAction))]
        public void FormatSingleExistingMember(int indentionLevel)
        {
            var indention = string.Concat(Enumerable.Repeat(" ", indentionLevel));

            string inputCode =
$@"
Option Explicit

Private mTest As Long

Private Type TestType
{indention}FirstValue As String
End Type
";
            var expectedUDT =
$@"
Private Type TestType
{indention}FirstValue As String
{indention}Test As Long
End Type
";

            var results = ExecuteTest(inputCode, "TestType", ("mTest", "Test"));
            StringAssert.Contains(expectedUDT, results);
        }

        [TestCase(4)]
        [TestCase(2)]
        [Category("Refactorings")]
        [Category(nameof(DeclareFieldsAsUDTMembersRefactoringAction))]
        public void FormatMatchesLastMemberIndent(int indentionLevel)
        {
            var indention = string.Concat(Enumerable.Repeat(" ", indentionLevel));
            var indentionFirstMember = string.Concat(Enumerable.Repeat(" ", 10));

            string inputCode =
$@"
Option Explicit

Private mTest As Long

Private Type TestType
{indentionFirstMember}FirstValue As String
{indention}SecondValue As Double
End Type
";
            var expectedUDT =
$@"
Private Type TestType
{indentionFirstMember}FirstValue As String
{indention}SecondValue As Double
{indention}Test As Long
End Type
";

            var results = ExecuteTest(inputCode, "TestType", ("mTest", "Test"));
            StringAssert.Contains(expectedUDT, results);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeclareFieldsAsUDTMembersRefactoringAction))]
        public void FormatPreservesComment()
        {
            var indention = string.Concat(Enumerable.Repeat(" ", 2));

            string inputCode =
$@"
Option Explicit

Private mTest As Long

Private Type TestType
{indention}FirstValue As String
{indention}SecondValue As Double    'This is a comment
End Type
";
            var expectedUDT =
$@"
Private Type TestType
{indention}FirstValue As String
{indention}SecondValue As Double    'This is a comment
{indention}Test As Long
End Type
";

            var results = ExecuteTest(inputCode, "TestType", ("mTest", "Test"));
            StringAssert.Contains(expectedUDT, results);
        }

        private string ExecuteTest(string inputCode, string udtIdentifier, params (string, string)[] fieldConversions) 
        {
            return RefactoredCode(inputCode, state => TestModel(state, udtIdentifier, fieldConversions));
        }

        private DeclareFieldsAsUDTMembersModel TestModel(RubberduckParserState state, string udtIdentifier, params (string fieldID, string udtMemberID)[] fieldConversions)
        {
            var udtDeclaration = GetUniquelyNamedDeclaration(state, DeclarationType.UserDefinedType, udtIdentifier);
            var conversions = new List<(VariableDeclaration field, string udtMemberID)>();
            foreach (var (fieldID, udtMemberID) in fieldConversions)
            {
                var fieldDeclaration = GetUniquelyNamedDeclaration(state, DeclarationType.Variable, fieldID) as VariableDeclaration;
                conversions.Add((fieldDeclaration, udtMemberID));
            }

            var model = new DeclareFieldsAsUDTMembersModel();
            foreach ((VariableDeclaration field, string udtMemberID) in conversions)
            {
                model.AssignFieldToUserDefinedType(udtDeclaration, field, udtMemberID);
            }
            return model;
        }

        private static Declaration GetUniquelyNamedDeclaration(IDeclarationFinderProvider declarationFinderProvider, DeclarationType declarationType, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(declarationType).Single(d => d.IdentifierName.Equals(identifier));
        }

        protected override IRefactoringAction<DeclareFieldsAsUDTMembersModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new DeclareFieldsAsUDTMembersRefactoringAction(state, rewritingManager, new CodeBuilder());
        }
    }
}
