using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField.EncapsulateFieldUseBackingField
{
    [TestFixture]
    public class EncapsulateFieldUseBackingFieldRefactoringActionTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase(false, "Name")]
        [TestCase(true, "Name")]
        [TestCase(false, null)]
        [TestCase(true, null)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void EncapsulatePublicField(bool isReadOnly, string propertyIdentifier)
        {
            var target = "fizz";
            var inputCode = $"Public {target} As Integer";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.DeclarationFinder.MatchName(target).Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, isReadOnly, propertyIdentifier);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create( new List<FieldEncapsulationModel>() { fieldModel });
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            var resultPropertyIdentifier = propertyIdentifier ?? target.CapitalizeFirstLetter();

            var backingField = propertyIdentifier != null
                ? target
                : $"{target}1";

            StringAssert.Contains($"Public Property Get {resultPropertyIdentifier}()", refactoredCode);
            StringAssert.Contains($"{resultPropertyIdentifier} = {backingField}", refactoredCode);

            if (isReadOnly)
            {
                StringAssert.DoesNotContain($"Public Property Let {resultPropertyIdentifier}(", refactoredCode);
                StringAssert.DoesNotContain($"{backingField} = ", refactoredCode);
            }
            else
            {
                StringAssert.Contains($"Public Property Let {resultPropertyIdentifier}(", refactoredCode);
                StringAssert.Contains($"{backingField} = ", refactoredCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void EmptyTargetSet()
        {
            var target = "fizz";
            var inputCode = $"Public {target} As Integer";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create(Enumerable.Empty<FieldEncapsulationModel>());
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);
            Assert.AreEqual(refactoredCode, inputCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void LocalReadonlyFieldsWriteReferencesLeftAsIs()
        {
            var target = "mStuff";
            var inputCode =
@"
Option Explicit

Public mStuff As Collection

Private Sub Class_Initialize()
    Set mStuff = New Collection
End Sub
";
            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.DeclarationFinder.MatchName(target).Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, true);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.Contains($"Set mStuff = New Collection", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void LocalReadonlyUDTMembersWriteReferencesLeftAsIs()
        {
            var target = "this";
            var inputCode =
@"
Option Explicit

Private Type TTest
    FirstVal As Long
    SecondVal As String
End Type

Private this As TTest

Private Sub Class_Initialize()
    this.FirstVal = 7
End Sub
";
            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.DeclarationFinder.MatchName(target).Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, true);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.Contains($"this.FirstVal = 7", refactoredCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        //Forces the 'Read Only" flag to false if there are external write references.
        //If invoked via the refactoring dialog, the user will not given the option to choose a Readonly implementation
        //if there are external write references.
        public void ExternallyReferencedField_IgnoresReadOnlyFlagIfExternalWriteReferencesExist(bool isReadOnly)
        {
            var propertyIdentifier = "Name";
            var codeClass1 =
@"Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            var codeClass2 =
@"Sub Foo()
    Dim theClass As Class1
    Set theClass = new Class1
    theClass.fizz = 0
    Bar theClass.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";
            (string name, string content, ComponentType compType) class1 = ("Class1", codeClass1, ComponentType.ClassModule);
            (string name, string content, ComponentType compType) class2 = ("Class2", codeClass2, ComponentType.ClassModule);

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {

                var field = state.DeclarationFinder.MatchName("fizz").Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, isReadOnly, propertyIdentifier);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
            }

            var vbe = MockVbeBuilder.BuildFromModules((class1.name, class1.content, class1.compType), (class2.name, class2.content, class2.compType)).Object;
            var refactoredCode = Support.RefactoredCode<EncapsulateFieldUseBackingFieldRefactoringAction, EncapsulateFieldUseBackingFieldModel>(vbe, modelBuilder);

            StringAssert.Contains("fizz = RHS", refactoredCode["Class1"]);
            StringAssert.Contains("Name = 1", refactoredCode["Class1"]);
            StringAssert.Contains($"theClass.{propertyIdentifier} = 0", refactoredCode["Class2"]);
            StringAssert.Contains($"Bar theClass.{propertyIdentifier}", refactoredCode["Class2"]);
            StringAssert.DoesNotContain("fizz", refactoredCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void EncapsulatePublicFields_DeeperNestedPathForPrivateUDTFieldReadonlyFlag()
        {
            var inputCode =
$@"
Option Explicit

Private Type PType1
    FirstValType1 As Long
    SecondValType1 As String
End Type


Private Type PType2
    FirstValType2 As Long
    SecondValType2 As String
    Third As PType1
End Type

Private mTypesField As PType2

Private Sub Class_Initialize()
    mTypesField.Third.SecondValType1 = ""Wah""
End Sub

Private Sub TestSub2()
    TestSub3 mTypesField.Third.SecondValType1
End Sub

Private Sub TestSub3(ByVal arg As String)
End Sub
";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mTypesField"));
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, false);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                var model = modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
                foreach (var candidate in model.EncapsulationCandidates)
                {
                    candidate.EncapsulateFlag = true;
                    candidate.IsReadOnly = true;
                }
                return model;
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.Contains("TypesField.Third.SecondValType1 = \"Wah\"", refactoredCode);
            StringAssert.Contains("TestSub3 SecondValType1", refactoredCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void RespectsGroupRelatedPropertiesIndenterSetting(bool groupRelatedProperties)
        {
            var inputCode =
@"
Public mTestField As Long
Public mTestField1 As Long
Public mTestField2 As Long
";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mTestField"));
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, false);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                var model = modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
                foreach (var candidate in model.EncapsulationCandidates)
                {
                    candidate.EncapsulateFlag = true;
                }
                return model;
            }

            var testIndenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                s.GroupRelatedProperties = groupRelatedProperties;
                return s;
            });

            var refactoredCode = RefactoredCode(inputCode, modelBuilder, testIndenter);

            var lines = refactoredCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var expectedGrouped = new[]
            {
                "Public Property Get TestField() As Long",
                "TestField = mTestField",
                "End Property",
                "Public Property Let TestField(ByVal RHS As Long)",
                "mTestField = RHS",
                "End Property",
                "",
                "Public Property Get TestField1() As Long",
                "TestField1 = mTestField1",
                "End Property",
                "Public Property Let TestField1(ByVal RHS As Long)",
                "mTestField1 = RHS",
                "End Property",
                "",
                "Public Property Get TestField2() As Long",
                "TestField2 = mTestField2",
                "End Property",
                "Public Property Let TestField2(ByVal RHS As Long)",
                "mTestField2 = RHS",
                "End Property",
                "",
            };

            var expectedNotGrouped = new[]
            {
                "Public Property Get TestField() As Long",
                "TestField = mTestField",
                "End Property",
                "",
                "Public Property Let TestField(ByVal RHS As Long)",
                "mTestField = RHS",
                "End Property",
                "",
                "Public Property Get TestField1() As Long",
                "TestField1 = mTestField1",
                "End Property",
                "",
                "Public Property Let TestField1(ByVal RHS As Long)",
                "mTestField1 = RHS",
                "End Property",
                "",
                "Public Property Get TestField2() As Long",
                "TestField2 = mTestField2",
                "End Property",
                "",
                "Public Property Let TestField2(ByVal RHS As Long)",
                "mTestField2 = RHS",
                "End Property",
                "",
            };

            var idx = 0;

            IList<string> expected = groupRelatedProperties
                ? expectedGrouped.ToList()
                : expectedNotGrouped.ToList();

            var refactoredLinesOfInterest = lines.SkipWhile(rl => !rl.Contains(expected[0]));

            Assert.IsTrue(refactoredLinesOfInterest.Any());

            foreach (var line in refactoredLinesOfInterest)
            {
                if (!line.Contains("="))
                {
                    StringAssert.AreEqualIgnoringCase(expected[idx], line);
                }
                idx++;
            }
        }

        private string RefactoredCode(string inputCode, Func<RubberduckParserState, EncapsulateFieldTestsResolver, EncapsulateFieldUseBackingFieldModel> modelBuilder, IIndenter indenter = null)
        {
            return Support.RefactoredCode<EncapsulateFieldUseBackingFieldRefactoringAction, EncapsulateFieldUseBackingFieldModel>(inputCode, modelBuilder, indenter);
        }
    }
}
