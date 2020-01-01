using NUnit.Framework;
using Moq;
using Rubberduck.Refactorings.EncapsulateField;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulationIdentifiersTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

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
            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, "fizz");
            Assert.IsTrue(encapsulatedField.TryValidateEncapsulationAttributes(out _));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameValuesPerSequenceOfPropertyNameChanges()
        {
            string inputCode = "Public fizz As String";

            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, "fizz");
            StringAssert.AreEqualIgnoringCase("fizz_1", encapsulatedField.FieldIdentifier);

            encapsulatedField.PropertyName = "Test";
            StringAssert.AreEqualIgnoringCase("fizz", encapsulatedField.FieldIdentifier);

            encapsulatedField.PropertyName = "Fizz";
            StringAssert.AreEqualIgnoringCase("fizz_1", encapsulatedField.FieldIdentifier);

            encapsulatedField.PropertyName = "Fiz";
            StringAssert.AreEqualIgnoringCase("fizz", encapsulatedField.FieldIdentifier);

            encapsulatedField.PropertyName = "Fizz";
            StringAssert.AreEqualIgnoringCase("fizz_1", encapsulatedField.FieldIdentifier);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameChange_UniqueParamName()
        {
            string inputCode = "Public value As String";

            var encapsulatedField = Support.RetrieveEncapsulateFieldCandidate(inputCode, "value");

            encapsulatedField.PropertyName = "Test";
            StringAssert.AreEqualIgnoringCase("value_1", encapsulatedField.ParameterName);

            encapsulatedField.PropertyName = "Value";
            StringAssert.AreEqualIgnoringCase("value_2", encapsulatedField.ParameterName);
        }

        [TestCase("strValue", "Value", "strValue")]
        [TestCase("m_Text", "Text", "m_Text")]
        [TestCase("notAHungarianName", "NotAHungarianName", "notAHungarianName_1")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void AccountsForHungarianNamesAndMemberPrefix(string inputName, string expectedPropertyName, string expectedFieldName)
        {
            var sut = new EncapsulationIdentifiers(inputName, (string name) => true);

            Assert.AreEqual(expectedPropertyName, sut.DefaultPropertyName);
            Assert.AreEqual(expectedFieldName, sut.DefaultNewFieldName);
        }
    }
}
