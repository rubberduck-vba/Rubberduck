using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.CreateUDTMember;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.CreateUDTMember
{
    [TestFixture]
    public class CreateUDTMemberRefactoringActionTests : RefactoringActionTestBase<CreateUDTMemberModel>
    {
        [TestCase(4)]
        [TestCase(2)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(CreateUDTMemberRefactoringAction))]
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
        [Category("Encapsulate Field")]
        [Category(nameof(CreateUDTMemberRefactoringAction))]
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
        [Category("Encapsulate Field")]
        [Category(nameof(CreateUDTMemberRefactoringAction))]
        public void FormatPreservesComments()
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

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(CreateUDTMemberRefactoringAction))]
        public void FormatMultipleInsertions()
        {
            string inputCode =
$@"
Option Explicit

Private mTest As Long
Private mTest1 As Long
Private mTest2 As Long

Private Type TestType
    FirstValue As String
    SecondValue As Double
End Type
";
            var expectedUDT =
$@"
Private Type TestType
    FirstValue As String
    SecondValue As Double
    Test As Long
    Test1 As Long
    Test2 As Long
End Type
";

            var results = ExecuteTest(inputCode, "TestType", ("mTest", "Test"), ("mTest1", "Test1"), ("mTest2", "Test2"));
            StringAssert.Contains(expectedUDT, results);
        }

        private string ExecuteTest(string inputCode, string udtIdentifier, params (string, string)[] fieldConversions) 
        {
            return RefactoredCode(inputCode, state => TestModel(state, udtIdentifier, fieldConversions));
        }

        private CreateUDTMemberModel TestModel(RubberduckParserState state, string udtIdentifier, params (string fieldID, string udtMemberID)[] fieldConversions)
        {
            var udtDeclaration = GetUniquelyNamedDeclaration(state, DeclarationType.UserDefinedType, udtIdentifier);

            var conversionPairs = new List<(Declaration, string)>();
            foreach (var (fieldID, udtMemberID) in fieldConversions)
            {
                var fieldDeclaration = GetUniquelyNamedDeclaration(state, DeclarationType.Variable, fieldID) as VariableDeclaration;
                conversionPairs.Add((fieldDeclaration, udtMemberID));
            }

            var model = new CreateUDTMemberModel(udtDeclaration, conversionPairs.ToArray());

            return model;
        }

        private static Declaration GetUniquelyNamedDeclaration(IDeclarationFinderProvider declarationFinderProvider, DeclarationType declarationType, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(declarationType).Single(d => d.IdentifierName.Equals(identifier));
        }

        protected override IRefactoringAction<CreateUDTMemberModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new CreateUDTMemberRefactoringAction(state, rewritingManager, CreateCodeBuilder());
        }

        private static ICodeBuilder CreateCodeBuilder()
            => new CodeBuilder(new Indenter(null, CreateIndenterSettings));

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }
    }
}
