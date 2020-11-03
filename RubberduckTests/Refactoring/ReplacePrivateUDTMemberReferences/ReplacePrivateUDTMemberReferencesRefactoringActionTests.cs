using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using RubberduckTests.Mocks;
using RubberduckTests.Refactoring;
using System;
using System.Linq;

namespace RubberduckTests.ReplacePrivateUDTMemberReferences
{
    [TestFixture]
    public class ReplacePrivateUDTMemberReferencesRefactoringActionTests : RefactoringActionTestBase<ReplacePrivateUDTMemberReferencesModel>
    {
        [TestCase("TheFirst", "TheSecond")]
        [TestCase("afirst", "asecond")] //respects casing
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
        public void RenameFieldReferences(string firstValueRefReplacement, string secondValueRefReplacement)
        {
            string inputCode =
$@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Sub Fizz(newValue As String)
    myBazz.FirstValue = newValue
End Sub

Public Sub Bazz(newValue As String)
    myBazz.SecondValue = newValue
End Sub
";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var testParam1 = new PrivateUDTExpressions("myBazz", "FirstValue")
            {
                InternalName = firstValueRefReplacement,
            };
            var testParam2 = new PrivateUDTExpressions("myBazz", "SecondValue")
            {
                InternalName = secondValueRefReplacement,
            };

            var results = RefactoredCode(vbe.Object, state => TestModel(state, false, testParam1, testParam2));
            StringAssert.Contains($"  {firstValueRefReplacement} = newValue", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains($"  {secondValueRefReplacement} = newValue", results[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
        public void RenameFieldReferences_WithMemberAccess()
        {
            string inputCode =
$@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Sub Fizz(newValue As String)
    With myBazz
        .FirstValue = newValue
    End With
End Sub

Public Sub Bazz(newValue As String)
    With myBazz
        .SecondValue = newValue
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var testParam1 = new PrivateUDTExpressions("myBazz", "FirstValue")
            {
                InternalName = "TheFirst",
            };
            var testParam2 = new PrivateUDTExpressions("myBazz", "SecondValue")
            {
                InternalName = "TheSecond",
            };

            var results = RefactoredCode(vbe.Object, state => TestModel(state, false, testParam1, testParam2));
            StringAssert.Contains($"  With myBazz{Environment.NewLine}", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("  TheFirst = newValue", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("  TheSecond = newValue", results[MockVbeBuilder.TestModuleName]);
        }

        private ReplacePrivateUDTMemberReferencesModel TestModel(RubberduckParserState state, bool moduleQualify = true, params PrivateUDTExpressions[] fieldConversions)
        {
            var fields = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false)
                .Select(v => v as VariableDeclaration);

            var factory =  new ReplacePrivateUDTMemberReferencesModelFactory(state);

            var model = factory.Create(fields);

            foreach (var fieldConversion in fieldConversions)
            {
                var fieldDeclaration = fields.Single(f => f.IdentifierName == fieldConversion.FieldID);
                var udtMember = model.UDTMembers
                    .Single(udtm => udtm.ParentDeclaration == fieldDeclaration.AsTypeDeclaration 
                        && udtm.IdentifierName == fieldConversion.UDTMemberID);

                var expressions = new PrivateUDTMemberReferenceReplacementExpressions(fieldConversion.InternalName);

                model.AssignUDTMemberReferenceExpressions(fieldDeclaration as VariableDeclaration, udtMember, expressions);
            }
            return model;
        }

        private static bool IsExternalReference(IdentifierReference identifierReference)
            => identifierReference.QualifiedModuleName != identifierReference.Declaration.QualifiedModuleName;

        private static Declaration GetUniquelyNamedDeclaration(IDeclarationFinderProvider declarationFinderProvider, DeclarationType declarationType, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(declarationType).Single(d => d.IdentifierName.Equals(identifier));
        }

        protected override IRefactoringAction<ReplacePrivateUDTMemberReferencesModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new ReplacePrivateUDTMemberReferencesRefactoringAction(rewritingManager);
        }

        private struct PrivateUDTExpressions
        {
            public PrivateUDTExpressions(string fieldID, string udtMemberIdentifier)
            {
                FieldID = fieldID;
                UDTMemberID = udtMemberIdentifier;
                _externalName = null;
                _internalName = null;
            }

            public string FieldID { set; get; }

            public string UDTMemberID {set; get;}

            private string _internalName;
            public string InternalName
            {
                set => _internalName = value;
                get => _internalName ?? FieldID.CapitalizeFirstLetter();
            }

            private string _externalName;
            public string ExternalName
            {
                set => _externalName = value;
                get => _externalName ?? FieldID.CapitalizeFirstLetter();
            }
    }
}
}
