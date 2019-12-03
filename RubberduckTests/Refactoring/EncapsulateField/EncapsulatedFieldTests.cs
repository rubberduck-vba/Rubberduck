using NUnit.Framework;
using Moq;
using Rubberduck.Refactorings.EncapsulateField;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulatedFieldTests
    {
        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameAttributeValidations_Internal(string originalFieldName, string newPropertyName, bool expectedResult)
        {
            string inputCode =
$@"Public {originalFieldName} As String";

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
            Assert.IsTrue(encapsulatedField.HasValidEncapsulationAttributes);
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

            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "First", DeclarationType.UserDefinedTypeMember);
            StringAssert.AreNotEqualIgnoringCase("First", encapsulatedField.PropertyName);
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
            Assert.IsTrue(encapsulatedField.HasValidEncapsulationAttributes);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameValuesPerSequenceOfPropertyNameChanges()
        {
            string inputCode = "Public fizz As String";

            var encapsulatedField = RetrieveEncapsulatedField(inputCode, "fizz");
            StringAssert.AreEqualIgnoringCase("fizz1", encapsulatedField.NewFieldName);

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

            var mock = new Mock<IEncapsulateFieldCandidate>();
            mock.SetupGet(m => m.AsTypeName).Returns("String");
            mock.SetupGet(m => m.PropertyName).Returns("Fizz");

            var newUserDefinedType = new UDTDeclarationGenerator("This_Type");
            newUserDefinedType.AddMember(mock.Object);

            var result = newUserDefinedType.TypeDeclarationBlock(CreateIndenter());
            StringAssert.Contains("Private Type This_Type", result);
            StringAssert.Contains(" Fizz As String", result);
            StringAssert.Contains("End Type", result);
        }

        private IEncapsulateFieldCandidate RetrieveEncapsulatedField(string inputCode, string fieldName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var match = state.DeclarationFinder.MatchName(fieldName).Single();
                return new EncapsulateFieldCandidate(match, new EncapsulateFieldNamesValidator(state)) as IEncapsulateFieldCandidate;
            }
        }

        private IEncapsulateFieldCandidate RetrieveEncapsulatedField(string inputCode, string fieldName, DeclarationType declarationType)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var match = state.DeclarationFinder.MatchName(fieldName).Where(m => m.DeclarationType.Equals(declarationType)).Single();
                return new EncapsulateFieldCandidate(match, new EncapsulateFieldNamesValidator(state)) as IEncapsulateFieldCandidate;
            }
        }

        private static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
    }
}
