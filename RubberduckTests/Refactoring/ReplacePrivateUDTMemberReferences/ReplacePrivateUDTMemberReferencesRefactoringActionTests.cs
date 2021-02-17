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
        [TestCase("afirst", "asecond")] //check for retention of casing
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
        public void ReplaceUdtMemberReferences(string firstValueRefReplacement, string secondValueRefReplacement)
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

            var testParam1 = (("myBazz", "FirstValue"), firstValueRefReplacement);
            var testParam2 = (("myBazz", "SecondValue"), secondValueRefReplacement);

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

            var testParam1 = (("myBazz", "FirstValue"), "TheFirst");
            var testParam2 = (("myBazz", "SecondValue"), "TheSecond");

            var results = RefactoredCode(vbe.Object, state => TestModel(state, false, testParam1, testParam2));
            StringAssert.Contains($"  With myBazz{Environment.NewLine}", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("  TheFirst = newValue", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("  TheSecond = newValue", results[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
        public void ReplaceAccessorExpression()
        {
            string inputCode =
$@"

Public exposed As String

Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Function GetTheFirstValue() As String
    GetTheFirstValue = myBazz.FirstValue
End Function

Public Sub SetTheFirstValue(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));

            var testParam1 = (("myBazz", "FirstValue"), "exposed");

            var results = RefactoredCode(vbe.Object, state => TestModel(state, false, testParam1));
            StringAssert.Contains("exposed = arg", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("GetTheFirstValue = exposed", results[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(ReplacePrivateUDTMemberReferencesRefactoringAction))]
        public void ReplacePublicUDTAccessorExpression()
        {
            string inputCode =
$@"

Public Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Public myBazz As TBazz

Public Function GetTheFirstValue() As String
    GetTheFirstValue = myBazz.FirstValue
End Function

Public Sub SetTheFirstValue(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            string referencingCode =
$@"

Public Function GetBazzFirst() As String
    GetBazzFirst = myBazz.FirstValue
End Function

Public Sub SetBazzFirst(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode),
                ("ReferencingModule", referencingCode));

            var testParam1 = (("myBazz", "FirstValue"), "NewProperty");

            var results = RefactoredCode(vbe.Object, state => TestModel(state, false, testParam1));
            StringAssert.Contains("NewProperty = arg", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("GetTheFirstValue = NewProperty", results[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains($"GetBazzFirst = {MockVbeBuilder.TestModuleName}.NewProperty", results["ReferencingModule"]);
            StringAssert.Contains($"{MockVbeBuilder.TestModuleName}.NewProperty = arg", results["ReferencingModule"]);
        }

        private ReplacePrivateUDTMemberReferencesModel TestModel(RubberduckParserState state, 
            bool moduleQualify = true, 
            params ((string instanceID, string udtMemberID) targetTuple, string replacementExpression)[] fieldConversions)
        {
            var fields = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(v => fieldConversions.Select(f => f.targetTuple.instanceID).Contains(v.IdentifierName))
                .Select(v => v as VariableDeclaration);

            var factory = new ReplacePrivateUDTMemberReferencesModelFactory(state);

            var model = factory.Create(fields);

            foreach (var field in fields)
            {
                var udtInstance = model.UserDefinedTypeInstance(field);

                foreach (var rf in udtInstance.UDTMemberReferences)
                {
                    var replacementExpression = fieldConversions.Where(f => f.targetTuple.instanceID == udtInstance.InstanceField.IdentifierName
                        && f.targetTuple.udtMemberID == rf.Declaration.IdentifierName)
                        .Select(f => f.replacementExpression).Single();
                    model.RegisterReferenceReplacementExpression(rf, replacementExpression);
                }
            }
            return model;
        }

        protected override IRefactoringAction<ReplacePrivateUDTMemberReferencesModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new ReplacePrivateUDTMemberReferencesRefactoringAction(rewritingManager);
        }
    }
}
