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
    public class EncapsulatedFieldTests
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
            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "fizz");
            Assert.IsTrue(encapsulatedField.HasValidEncapsulationAttributes);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameValuesPerSequenceOfPropertyNameChanges()
        {
            string inputCode = "Public fizz As String";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "fizz");
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

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "value");

            encapsulatedField.PropertyName = "Test";
            StringAssert.AreEqualIgnoringCase("value_value", encapsulatedField.ParameterName);

            encapsulatedField.PropertyName = "Value";
            StringAssert.AreEqualIgnoringCase("Value_value_1_value", encapsulatedField.ParameterName);
        }
    }
}
