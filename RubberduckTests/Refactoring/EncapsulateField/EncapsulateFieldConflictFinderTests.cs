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
        public void RespectsMemberAccessExpressionUsingMe()
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
        public void RespectsWithMemberAccessExpressionUsingMe()
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

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldConflictFinder))]
        public void RespectsWithMemberAccessExpressionUsingMe_Parameter()
        {
            var targetField = "Fizz";
            var inputCode =
$@"Public Fizz As Integer

Public Sub CopyFromAnotherInstance(Fizz As MockVbeBuilder.TestModuleName)
    Me.Fizz = Fizz.Fizz
End Sub";

            var result = TestIsConflictingIdentifier("Fizz", inputCode, targetField);
            Assert.IsFalse(result);
        }

        [TestCase("Dim thisConflicts As Long", null, true)]
        [TestCase("Dim thisConflicts As Long", "Me.", false)]
        [TestCase("Const thisConflicts As Long = 10", null, true)]
        [TestCase("Const thisConflicts As Long = 10", "Me.", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldConflictFinder))]
        public void LocalDeclarationsRespectsQualifier(string declaration, string qualifier, bool expected)
        {
            var targetField = "Fizz";
            var inputCode =
$@"Public Fizz As Integer

Public Sub DoSomething()
    {declaration}
    {qualifier}Fizz = thisConflicts + 1
End Sub";

            var result = TestIsConflictingIdentifier("thisConflicts", inputCode, targetField);
            Assert.AreEqual(expected, result);
        }

        [TestCase("thisConFlicts", true)]
        [TestCase("DoSomething", true)]
        [TestCase("TestString", true)]
        [TestCase("TestConst", true)]
        [TestCase("FirstValue", false)]
        [TestCase("LocalConst", false)]
        [TestCase("LocalVariable", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldConflictFinder))]
        public void DetectsModuleEntityConflicts(string newName, bool expected)
        {
            var targetField = "Fizz";
            var inputCode =
$@"
Public Fizz As Integer
Private TestString As String

Private Const TestConst As Long = 10

Private Enum TestEnum
    ThisConflicts
    EnumTwo
End Enum

Private Type TestType
    FirstValue As Long
End Type

Private Sub DoSomething()
    Const localConst As String = ""Test""
    Dim localVariable As Long
End Sub
";

            var result = TestIsConflictingIdentifier(newName, inputCode, targetField);
            Assert.AreEqual(expected, result);
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
