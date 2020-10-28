using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldConflictFinderTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [SetUp]
        public void ExecutesBeforeAllTests()
        {
            Support.ResetResolver();
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldConflictFinder))]
        public void AvoidsFalsePositive_MemberAccessExpression_Me()
        {
            var targetField = "Fizz";

            var inputCode =
$@"Public {targetField} As Integer

Public Sub NoConflict(Fizz As Integer)
    Me.Fizz = Fizz
End Sub";

            var result = TestIsConflictingIdentifier("Fizz", inputCode, targetField);
            Assert.IsFalse(result);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldConflictFinder))]
        public void AvoidsFalsePositive_WithMemberAccessExpression_Me()
        {
            var targetField = "Fizz";

            var inputCode =
$@"Public Fizz As Integer

Public Sub NoConflict(Fizz As Integer)
    With Me
       .Fizz = Fizz
    End With
End Sub";
            var result = TestIsConflictingIdentifier("Fizz", inputCode, targetField);
            Assert.IsFalse(result);
        }

        private bool TestIsConflictingIdentifier(string testIdentifier, string inputCode, string targetFieldName)
        {
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.ClassModule));

            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var field = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Single(d => d.IdentifierName == targetFieldName);

                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration);

                var modelFactory = Support.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>(state);
                var model = modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });

                var efCandidate = model.EncapsulationCandidates.First(c => c.Declaration == field);
                efCandidate.EncapsulateFlag = true;

                return model.ConflictFinder.IsConflictingIdentifier(efCandidate, testIdentifier, out _);
            }
        }
    }
}
