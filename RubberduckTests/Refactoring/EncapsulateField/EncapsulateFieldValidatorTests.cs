using NUnit.Framework;
using Moq;
using Rubberduck.Refactorings.EncapsulateField;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;
using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Symbols;
using System;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldValidatorTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void VBAIdentifier_Property(string originalFieldName, string newPropertyName, bool expectedResult)
        {
            string inputCode =
$@"Public {originalFieldName} As String";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, originalFieldName);

            encapsulatedField.PropertyName = newPropertyName;
            var field = encapsulatedField as IEncapsulateFieldCandidateValidations;
            Assert.AreEqual(expectedResult, field.HasVBACompliantPropertyIdentifier);
        }

        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void VBAIdentifier_FieldName(string originalFieldName, string newFieldName, bool expectedResult)
        {
            string inputCode =
$@"Public {originalFieldName} As String";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, originalFieldName);

            encapsulatedField.FieldIdentifier = newFieldName;
            var field = encapsulatedField as IEncapsulateFieldCandidateValidations;
            Assert.AreEqual(expectedResult, field.HasVBACompliantFieldIdentifier);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_ReadOnlyRequiresSet()
        {
            const string inputCode =
                @"|Private fizz As Collection";

            const string expectedCode =
                @"Private fizz As Collection

Public Property Get Name() As Collection
    Set Name = fizz
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", isReadonly: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode.Trim(), actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameNotDuplicated()
        {
            const string inputCode =
                @"Public var|iable As Integer, variable1 As Long, variable2 As Integer";

            var userInput = new UserInputDataObject()
                .AddAttributeSet("variable")
                .AddAttributeSet("variable1")
                .AddAttributeSet("variable2");

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Variable() As Integer", actualCode);
            StringAssert.Contains("Variable = variable_1", actualCode);
            StringAssert.Contains("Public Property Get Variable1() As Long", actualCode);
            StringAssert.Contains("Variable1 = variable1_1", actualCode);
            StringAssert.Contains("Public Property Get Variable2() As Integer", actualCode);
            StringAssert.Contains("Variable2 = variable2_1", actualCode);
            StringAssert.DoesNotContain("Public Property Get Variable3() As Integer", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberPropertyDefaultsToValidValue()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public myBar As TBar

Private Function First() As String
    First = myBar.First
End Function";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "First", DeclarationType.UserDefinedTypeMember);
            var validation = encapsulatedField as IEncapsulateFieldCandidateValidations;
            var result = validation.HasConflictingPropertyIdentifier;
            Assert.AreEqual(true, validation.HasConflictingPropertyIdentifier);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameDefaultsToNonConflictName()
        {
            string inputCode =
$@"Public fizz As String

            Private fizzle As String

            'fizz1 is the initial default name for encapsulating 'fizz'            
            Public Property Get Fizz_1() As String
                Fizz_1 = fizzle
            End Property

            Public Property Let Fizz_1(ByVal value As String)
                fizzle = value
            End Property
            ";
            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "fizz");
            Assert.IsTrue(encapsulatedField.HasValidEncapsulationAttributes);
        }

        [TestCase("Name")]
        [TestCase("mName")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserEntersConflictingName(string userModifiedPropertyName)
        {
            string inputCode =
$@"Public fizz As String

            Private mName As String

            Public Property Get Name() As String
                Name = mName
            End Property

            Public Property Let Name(ByVal value As String)
                mName = value
            End Property
            ";

            var userInput = new UserInputDataObject("fizz", userModifiedPropertyName, true);

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, "fizz", DeclarationType.Variable, presenterAction, "fizz");

            Assert.IsFalse(model["fizz"].HasValidEncapsulationAttributes);
        }

        [TestCase("Number", "Bazzle", true, true)]
        [TestCase("Number", "Number", false, false)]
        [TestCase("Test", "Number", false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserModificationIsExistingPropertyNameConflicts(string fizz_modifiedPropertyName, string bazz_modifiedPropertyName, bool fizz_expectedResult, bool bazz_expectedResult)
        {
            string inputCode =
$@"Public fizz As Integer
Public bazz As Integer
Public buzz As Integer

Private mTest As Integer

Public Property Get Test() As Integer
    Test = mTest
End Property";

            var userInput = new UserInputDataObject()
                .AddAttributeSet("fizz", fizz_modifiedPropertyName, true)
                .AddAttributeSet("bazz", bazz_modifiedPropertyName, true);

            var presenterAction = Support.SetParameters(userInput);
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, "fizz", DeclarationType.Variable, presenterAction, "fizz", "bazz", "buzz");

            Assert.AreEqual(fizz_expectedResult, model["fizz"].HasValidEncapsulationAttributes, "fizz failed");
            Assert.AreEqual(bazz_expectedResult, model["bazz"].HasValidEncapsulationAttributes, "bazz failed");
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFields_UDTConflicts()
        {
            string inputCode =
$@"
Private Type TBar
    First As Long
    Second As String
End Type

Public this As TBar

Public that As TBar
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var thisField = CreateEncapsulatedFieldMock("this.First", "Long", vbe.SelectedVBComponent, "First");
                var thatField = CreateEncapsulatedFieldMock("that.First", "String", vbe.SelectedVBComponent, "First");

                var validator = new EncapsulateFieldNamesValidator(
                        state,
                        () => new List<IEncapsulateFieldCandidate>()
                            {
                                thisField,
                                thatField,
                            });
                validator.RegisterFieldCandidate(thisField);
                validator.RegisterFieldCandidate(thatField);

                Assert.IsTrue(validator.HasConflictingPropertyIdentifier(thisField));
                //Assert.Less(0, validator.HasNewPropertyNameConflicts(thisField, thisField.QualifiedModuleName, Enumerable.Empty<Declaration>()));
            }
        }
        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameConflictsWithModuleVariable()
        {
            string inputCode =
$@"
Public longValue As Long

Public wholeNumber As String
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            CreateAndParse(vbe, ThisTest);

            void ThisTest(IDeclarationFinderProvider declarationProviderProvider)
            {
                var wholeNumber = declarationProviderProvider.DeclarationFinder.MatchName("wholeNumber").Single();
                var longValue = declarationProviderProvider.DeclarationFinder.MatchName("longValue").Single();

                var fields = new List<IEncapsulateFieldCandidate>();

                var validator = new EncapsulateFieldNamesValidator(declarationProviderProvider, () => fields);

                var encapsulatedWholeNumber = new EncapsulateFieldCandidate(wholeNumber, validator);
                var encapsulatedLongValue = new EncapsulateFieldCandidate(longValue, validator);
                fields.Add(new EncapsulateFieldCandidate(wholeNumber, validator));
                fields.Add(new EncapsulateFieldCandidate(longValue, validator));

                encapsulatedWholeNumber.PropertyName = "LongValue";
                Assert.IsTrue(validator.HasConflictingPropertyIdentifier(encapsulatedWholeNumber));
                //Assert.Less(0, validator.HasNewPropertyNameConflicts(encapsulatedWholeNumber, encapsulatedWholeNumber.QualifiedModuleName, new Declaration[] { encapsulatedWholeNumber.Declaration }));
            }
        }

        private void CreateAndParse(IVBE vbe, Action<IDeclarationFinderProvider> theTest)
        {
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                theTest(state as IDeclarationFinderProvider);
            }
        }

        private IEncapsulateFieldCandidate CreateEncapsulatedFieldMock(string targetID, string asTypeName, IVBComponent component, string modifiedPropertyName = null, bool encapsulateFlag = true )
        {
            var identifiers = new EncapsulationIdentifiers(targetID);
            var attributesMock = CreateAttributesMock(targetID, asTypeName, modifiedPropertyName);

            var mock = new Mock<IEncapsulateFieldCandidate>();
            mock.SetupGet(m => m.TargetID).Returns(identifiers.TargetFieldName);
            mock.SetupGet(m => m.FieldIdentifier).Returns(identifiers.Field);
            mock.SetupGet(m => m.PropertyName).Returns(modifiedPropertyName ?? identifiers.Property);
            mock.SetupGet(m => m.AsTypeName).Returns(asTypeName);
            mock.SetupGet(m => m.EncapsulateFlag).Returns(encapsulateFlag);
            //mock.SetupGet(m => m.EncapsulationAttributes).Returns(attributesMock);
            mock.SetupGet(m => m.QualifiedModuleName).Returns(new QualifiedModuleName(component));
            return mock.Object;
        }

        private IEncapsulateFieldCandidate CreateAttributesMock(string targetID, string asTypeName, string modifiedPropertyName = null, bool encapsulateFlag = true)
        {
            var identifiers = new EncapsulationIdentifiers(targetID);
            var mock = new Mock<IEncapsulateFieldCandidate>();
            mock.SetupGet(m => m.IdentifierName).Returns(identifiers.TargetFieldName);
            mock.SetupGet(m => m.FieldIdentifier).Returns(identifiers.Field);
            mock.SetupGet(m => m.PropertyName).Returns(modifiedPropertyName ?? identifiers.Property);
            mock.SetupGet(m => m.AsTypeName).Returns(asTypeName);
            mock.SetupGet(m => m.EncapsulateFlag).Returns(encapsulateFlag);
            return mock.Object;
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }
}
