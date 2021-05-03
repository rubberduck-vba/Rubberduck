using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ModifyUserDefinedType;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.ModifyUserDefinedType
{
    [TestFixture]
    public class ModifyUserDefinedTypeRefactoringActionTests : RefactoringActionTestBase<ModifyUserDefinedTypeModel>
    {
        [TestCase("Private mTest As Long", DeclarationType.Variable)]
        [TestCase("Private Const mTest As Long = 10", DeclarationType.Constant)]
        [TestCase("Private Function mTest() As Long\r\nEnd Function", DeclarationType.Function)]
        [TestCase("Private Property Get mTest() As Long\r\nEnd Property", DeclarationType.PropertyGet)]
        [TestCase("Private Type ProtoType\r\n mTest As Long\r\nEnd Type", DeclarationType.UserDefinedTypeMember)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ModifyUserDefinedTypeRefactoringAction))]
        public void AddSingleMember(string prototypeDeclaration, DeclarationType declarationType)
        {
            var inputCode =
$@"
Option Explicit

Private Type TestType
    FirstValue As String
End Type

{prototypeDeclaration}
";
            var expectedUDT =
$@"
Private Type TestType
    FirstValue As String
    Test As Long
End Type
";

            ExecuteTest(inputCode, "TestType", expectedUDT, ("mTest", "Test", declarationType));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ModifyUserDefinedTypeRefactoringAction))]
        public void MultipleNewMembers()
        {
            var inputCode =
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

            ExecuteTest(inputCode, "TestType", expectedUDT, ("mTest", "Test", DeclarationType.Variable), ("mTest1", "Test1", DeclarationType.Variable), ("mTest2", "Test2", DeclarationType.Variable));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ModifyUserDefinedTypeRefactoringAction))]
        public void MultipleNewMembersRemoveMultiple()
        {
            var inputCode =
$@"
Option Explicit

Private mTest As Long
Private mTest1 As Long
Private mTest2 As Long

Private Type TestType
    FirstValue As String
    SecondValue As Double
    ThirdValue As Byte
End Type
";
            var expectedUDT =
$@"
Private Type TestType
    SecondValue As Double
    Test As Long
    Test1 As Long
    Test2 As Long
End Type
";
            var adds = new List<(string, string, DeclarationType)>()
            {
                ("mTest", "Test", DeclarationType.Variable),
                ("mTest1", "Test1", DeclarationType.Variable),
                ("mTest2", "Test2", DeclarationType.Variable)
            };

            var removes = new List<string>()
            {
                "FirstValue",
                "ThirdValue"
            };

            ExecuteTest(inputCode, "TestType", expectedUDT, adds, removes);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ModifyUserDefinedTypeRefactoringAction))]
        public void RemoveOnlyMultiple()
        {
            var inputCode =
$@"
Option Explicit

Private Type TestType
    FirstValue As String
    SecondValue As Double
    ThirdValue As Byte
    FourthValue As Integer
End Type
";
            var expectedUDT =
$@"
Private Type TestType
    FirstValue As String
    SecondValue As Double
End Type
";
            var removes = new List<string>()
            {
                "FourthValue",
                "ThirdValue"
            };

            ExecuteTest(inputCode, "TestType", expectedUDT, Enumerable.Empty<(string,string,DeclarationType)>(), removes);
        }

        private void ExecuteTest(string inputCode, string udtIdentifier, string expectedUDT, params (string, string, DeclarationType)[] fieldConversions)
        {
            ExecuteTest(inputCode, udtIdentifier, expectedUDT, (IEnumerable<(string, string, DeclarationType)>)fieldConversions);
        }

        private void ExecuteTest(string inputCode, string udtIdentifier, string expectedUDT, IEnumerable<(string, string, DeclarationType)> fieldConversions, IEnumerable<string> udtMemberIdentifiers = null)
        {
            var results = RefactoredCode(inputCode, state => TestModel(state, udtIdentifier, fieldConversions, udtMemberIdentifiers ?? Enumerable.Empty<string>()));

            var refactoredCode = results.Trim().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

            var refactored = refactoredCode.SkipWhile(r => !r.Contains("Private Type"));

            var expected = expectedUDT.Trim().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            for (var idx = 0; idx < expected.Count(); idx++)
            {
                //Remove Indenter formatting effects from refactoring results evaluation
                Assert.AreEqual(expected[idx].Trim(), refactored.ElementAt(idx).Trim());
            }
        }

        private ModifyUserDefinedTypeModel TestModel(RubberduckParserState state, string udtIdentifier, IEnumerable<(string fieldID, string udtMemberID, DeclarationType declarationType)> fieldConversions, IEnumerable<string> removals)
        {
            var udtDeclaration = GetUniquelyNamedDeclaration(state, DeclarationType.UserDefinedType, udtIdentifier);
            var model = new ModifyUserDefinedTypeModel(udtDeclaration);

            foreach (var (fieldID, udtMemberID, declarationType) in fieldConversions)
            {
                var fieldDeclaration = GetUniquelyNamedDeclaration(state, declarationType, fieldID);
                model.AddNewMemberPrototype(fieldDeclaration, udtMemberID);
            }

            foreach (var udtMemberIdentifier in removals)
            {
                var udtMember = GetUniquelyNamedDeclaration(state, DeclarationType.UserDefinedTypeMember, udtMemberIdentifier);
                model.RemoveMember(udtMember);
            }

            return model;
        }

        private static Declaration GetUniquelyNamedDeclaration(IDeclarationFinderProvider declarationFinderProvider, DeclarationType declarationType, string identifier) 
            => declarationFinderProvider.DeclarationFinder.UserDeclarations(declarationType).Single(d => d.IdentifierName.Equals(identifier));

        protected override IRefactoringAction<ModifyUserDefinedTypeModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new ModifyUserDefinedTypeRefactoringAction(state, rewritingManager, CreateCodeBuilder());
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
