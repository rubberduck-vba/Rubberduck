using NUnit.Framework;
using Rubberduck.Refactorings.EncapsulateField;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulationIdentifiersTests
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
        public void FieldNameAttributeValidation_DefaultsToAvailableFieldName()
        {
            var inputCode =
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
            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, "fizz");
            Assert.IsTrue(encapsulatedField.TryValidateEncapsulationAttributes(out _));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameValuesPerSequenceOfPropertyNameChanges()
        {
            var inputCode = "Public fizz As String";

            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, "fizz");
            StringAssert.AreEqualIgnoringCase("fizz1", encapsulatedField.BackingIdentifier);

            encapsulatedField.PropertyIdentifier = "Test";
            StringAssert.AreEqualIgnoringCase("fizz", encapsulatedField.BackingIdentifier);

            encapsulatedField.PropertyIdentifier = "Fizz";
            StringAssert.AreEqualIgnoringCase("fizz1", encapsulatedField.BackingIdentifier);

            encapsulatedField.PropertyIdentifier = "Fiz";
            StringAssert.AreEqualIgnoringCase("fizz", encapsulatedField.BackingIdentifier);

            encapsulatedField.PropertyIdentifier = "Fizz";
            StringAssert.AreEqualIgnoringCase("fizz1", encapsulatedField.BackingIdentifier);
        }

        [TestCase("strValue", "Value", "strValue")]
        [TestCase("m_Text", "Text", "m_Text")]
        [TestCase("notAHungarianName", "NotAHungarianName", "notAHungarianName1")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void AccountsForHungarianNamesAndMemberPrefix(string inputName, string expectedPropertyName, string expectedFieldName)
        {
            var sut = new EncapsulationIdentifiers(inputName);

            Assert.AreEqual(expectedPropertyName, sut.DefaultPropertyName);
            Assert.AreEqual(expectedFieldName, sut.DefaultNewFieldName);
        }
    }
}
