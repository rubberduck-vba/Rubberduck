using System;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulatedFieldTests //: InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameAttributeValidations_Internal(string originalFieldName, string newPropertyName, bool expectedResult)
        {
            string inputCode =
$@"Public {originalFieldName} As String";

            var selection = new Selection(1, 1);
            var encapsulatedField = RetrieveEncapsulatedField(inputCode, originalFieldName);

            encapsulatedField.PropertyName = newPropertyName;

            Assert.AreEqual(expectedResult, encapsulatedField.HasValidEncapsulationAttributes);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameAttributeValidation_DefaultsToAvailableFieldName()
        { 
            string inputCode =
$@"Public fizz As String

            'fizz1 is the intial default name for encapsulating 'fizz'            
            Private fizz1 As String

            Public Property Get Name() As String
                Name = fizz1
            End Property

            Public Property Let Name(ByVal value As String)
                fizz1 = value
            End Property
            ";
            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "fizz");
            Assert.AreEqual(true, encapsulatedField.HasValidEncapsulationAttributes);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameAttributeValidation_DefaultsToAvailablePropertyName()
        {
            string inputCode =
$@"Public fizz As String

            Private fizzle As String

            'fizz1 is the initial default name for encapsulating 'fizz'            
            Public Property Get Fizz1() As String
                Fizz1 = fizzle
            End Property

            Public Property Let Fizz1(ByVal value As String)
                fizzle = value
            End Property
            ";
            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "fizz");
            Assert.AreEqual(true, encapsulatedField.HasValidEncapsulationAttributes);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameChange_Multiple()
        {
            string inputCode = "Public fizz As String";

            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "fizz");

            encapsulatedField.PropertyName = "Test";
            StringAssert.AreEqualIgnoringCase("fizz", encapsulatedField.NewFieldName);

            encapsulatedField.PropertyName = "Fizz";
            StringAssert.AreEqualIgnoringCase("fizz1", encapsulatedField.NewFieldName);

            encapsulatedField.PropertyName = "Fiz";
            StringAssert.AreEqualIgnoringCase("fizz", encapsulatedField.NewFieldName);

            encapsulatedField.PropertyName = "Fizz";
            StringAssert.AreEqualIgnoringCase("fizz1", encapsulatedField.NewFieldName);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameChange_UniqueParamName()
        {
            string inputCode = "Public value As String";

            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "value");

            encapsulatedField.PropertyName = "Test";
            StringAssert.AreEqualIgnoringCase("value_value", encapsulatedField.EncapsulationAttributes.ParameterName);

            encapsulatedField.PropertyName = "Value";
            StringAssert.AreEqualIgnoringCase("Value_value1_value", encapsulatedField.EncapsulationAttributes.ParameterName);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void CreateUDT()
        {
            string inputCode =
$@"Public fizz As String";

            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "fizz");

            var udtTest = new EncapsulationUDT(CreateIndenter());
            udtTest.AddMember(encapsulatedField);
            var result = udtTest.TypeDeclarationBlock;
            StringAssert.Contains("Fizz As String", result);
        }

        private IEncapsulatedFieldDeclaration RetrieveEncapsulatedField(string inputCode, string fieldName)//, Selection selection) //, Func<TModel, TModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var module = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .Single(declaration => declaration.IdentifierName == selectedComponentName)
                    .QualifiedModuleName;

                var match = state.DeclarationFinder.MatchName(fieldName).Single();
                return new EncapsulatedFieldDeclaration(match, new EncapsulateFieldNamesValidator(state)) as IEncapsulatedFieldDeclaration;
            }
        }

        private ClientEncapsulationAttributes UserModifiedEncapsulationAttributes(string field, string property = null, bool? isReadonly = null, bool encapsulateFlag = true, string newFieldName = null)
        {
            var clientAttrs = new ClientEncapsulationAttributes(field);
            clientAttrs.NewFieldName = newFieldName ?? clientAttrs.NewFieldName;
            clientAttrs.PropertyName = property ?? clientAttrs.PropertyName;
            clientAttrs.ReadOnly = isReadonly ?? clientAttrs.ReadOnly;
            clientAttrs.EncapsulateFlag = encapsulateFlag;
            return clientAttrs;
        }

        private static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
    }
}
