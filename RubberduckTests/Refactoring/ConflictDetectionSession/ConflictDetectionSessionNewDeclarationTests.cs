using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using TestResolver = RubberduckTests.Refactoring.ConflictDetectionSessionTestsResolver;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ConflictDetectionSessionNewDeclarationTests
    {
        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [TestCase(DeclarationType.EnumerationMember)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithField(DeclarationType newDeclarationType)
        {
            var expectedName = "mTestVar1";
            var sourceCode =
$@"
Private mTestVar As Long
";
            var nonConflictName = RunNewDeclarationTest(("mTestVar", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [TestCase(DeclarationType.EnumerationMember)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithFieldConstant(DeclarationType newDeclarationType)
        {
            var expectedName = "mTestVar1";
            var sourceCode =
$@"
Private Const MTestVAR As Long = 453
";
            var nonConflictName = RunNewDeclarationTest(("mTestVar", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);

        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [TestCase(DeclarationType.EnumerationMember)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithFieldFunction(DeclarationType newDeclarationType)
        {
            var expectedName = "Fizz1";
            var sourceCode =
$@"
Private Function Fizz() As Long
End Function
";
            var nonConflictName = RunNewDeclarationTest(("Fizz", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [TestCase(DeclarationType.EnumerationMember)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithFieldSubroutine(DeclarationType newDeclarationType)
        {
            var expectedName = "Fazz1";
            var sourceCode =
$@"
Private Sub Fazz()
End Sub
";
            var nonConflictName = RunNewDeclarationTest(("Fazz", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [TestCase(DeclarationType.EnumerationMember)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithEnumMember(DeclarationType newDeclarationType)
        {
            var expectedName = "FirstValue1";
            var sourceCode =
$@"
Private Enum ETest
    FirstValue = 34
End Enum
";
            var nonConflictName = RunNewDeclarationTest(("FirstValue", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }


        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNamesRespectedByRename()
        {
            var expectedName = "SecondVariable1";
            var sourceCode =
$@"
Private Sub Fizz()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var stdModule = state.DeclarationFinder.DeclarationsWithType(DeclarationType.ProceduralModule)
                                .Where(d => d.IdentifierName.Equals(MockVbeBuilder.TestModuleName)).Single();

                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Procedure)
                                .Where(d => d.IdentifierName.Equals("Fizz")).Single();

                var conflictDetectionSession = TestResolver.Resolve<IConflictDetectionSessionFactory>(state).Create();
                foreach (var newVarName in new string[] {"FirstVariable", "SecondVariable"})
                    //conflictDetectionSession.NewDeclarationHasConflict(newVarName,
                    //                                    DeclarationType.Variable,
                    //                                    Accessibility.Private,
                    //                                    stdModule as ModuleDeclaration, stdModule,
                    //                                    out _);

                conflictDetectionSession.TryProposeNewDeclaration(newVarName,
                                                    DeclarationType.Variable,
                                                    Accessibility.Private,
                                                    stdModule as ModuleDeclaration, stdModule, 
                                                    out _,
                                                    false);

                //conflictDetectionSession.HasRenameConflict(target, "SecondVariable", out _); // var nonConflictName);
                conflictDetectionSession.TryProposeRenamePair(target, "SecondVariable"); // var nonConflictName);
                //var nonConflictName = conflictDetectionSession.GenerateNoConflictRename(target, "SecondVariable");
                StringAssert.AreEqualIgnoringCase(expectedName, conflictDetectionSession.ConflictFreeRenamePairs.Single(pr => pr.target == target).newName);
            }
        }

        [TestCase(MockVbeBuilder.TestModuleName)]
        [TestCase(MockVbeBuilder.TestProjectName)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleDeclarations(string newModuleName)
        {
            var expectedName = newModuleName == MockVbeBuilder.TestModuleName
                    ? IncrementIdentifier(MockVbeBuilder.TestModuleName)
                    : IncrementIdentifier(MockVbeBuilder.TestProjectName);

            var sourceCode =
$@"
Private mTest As Long
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var stdModule = state.DeclarationFinder.DeclarationsWithType(DeclarationType.ProceduralModule)
                                .Where(d => d.IdentifierName.Equals(MockVbeBuilder.TestModuleName)).Single();

                var namingToolsSession = TestResolver.Resolve<IConflictDetectionSessionFactory>(state).Create();
                namingToolsSession.NewModuleDeclarationHasConflict(newModuleName,
                                                    stdModule.ProjectId,
                                                    out var nonConflictName);

                StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
            }
        }

        private string RunNewDeclarationTest((string ID, DeclarationType Type, Accessibility accessibility) target, string sourceCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var destination = state.DeclarationFinder.DeclarationsWithType(DeclarationType.ProceduralModule)
                                .Where(d => d.IdentifierName.Equals(MockVbeBuilder.TestModuleName)).Single();

                var nameConflictManager = TestResolver.Resolve<IConflictDetectionSessionFactory>(state);
                var conflictSession = nameConflictManager.Create();
                //conflictSession.NewDeclarationHasConflict(target.ID,
                //                                        target.Type,
                //                                        target.accessibility,
                //                                        destination as ModuleDeclaration,
                //                                        destination,
                //                                        out var nonConflictName);
                conflictSession.TryProposeNewDeclaration(target.ID,
                                                        target.Type,
                                                        target.accessibility,
                                                        destination as ModuleDeclaration,
                                                        destination,
                                                        out int retrievalKey);

                //var results = new List<((string ID, DeclarationType decType, string ModuleName), string ResolvedID)>();
                //results.AddRange(conflictSession.NewDeclarationIdentifiers);

                //var theOne = results.Where(rID => rID.Item1.ID.Equals(target.ID)
                //                                && rID.Item1.decType.Equals(target.Type)
                //                                && rID.Item1.ModuleName.Equals(destination.IdentifierName)).First();
                (int key, string newName) = conflictSession.NewDeclarationIdentifiers.Single(pr => pr.keyID == retrievalKey);
                //var theOne = conflictSession.ConflictFreeRenamePairs.Single(pr => pr.GetHashCode() == retrievalKey);
                return newName;
            }
        }

        private static string IncrementIdentifier(string identifier)
        {
            var numeric = string.Concat(identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }
    }
}
