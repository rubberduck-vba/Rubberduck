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

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldValidatorTests
    {
        //private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("Number")]
        [TestCase("Test")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFields_PropertyNameConflicts(string modifiedPropertyName)
        {
            string inputCode =
$@"Public fizz As Integer
Public bazz As Integer
Public buzz As Integer
Private mTest As Integer

Public Property Get Test() As Integer
    Test = mTest
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var mockFizz = CreateEncapsulatedFieldMock("fizz", "Integer", vbe.SelectedVBComponent, modifiedPropertyName);
                var mockBazz = CreateEncapsulatedFieldMock("bazz", "Integer", vbe.SelectedVBComponent, "Whole");
                var mockBuzz = CreateEncapsulatedFieldMock("buzz", "Integer", vbe.SelectedVBComponent, modifiedPropertyName);

                var validator = new EncapsulateFieldNamesValidator(
                        state,
                        () => new List<IEncapsulatedFieldDeclaration>()
                            {
                                mockFizz,
                                mockBazz,
                                mockBuzz
                            });

                Assert.IsTrue(validator.HasNewPropertyNameConflicts(mockFizz.EncapsulationAttributes, mockFizz.QualifiedModuleName, (Declaration dec) => false));
            }
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
            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var thisField = CreateEncapsulatedFieldMock("this.First", "Long", vbe.SelectedVBComponent, "First");
                var thatField = CreateEncapsulatedFieldMock("that.First", "String", vbe.SelectedVBComponent, "First");

                var validator = new EncapsulateFieldNamesValidator(
                        state,
                        () => new List<IEncapsulatedFieldDeclaration>()
                            {
                                thisField,
                                thatField,
                            });


                Assert.IsTrue(validator.HasNewPropertyNameConflicts(thisField.EncapsulationAttributes, thisField.QualifiedModuleName, (Declaration dec) => false));
            }
        }

        private IEncapsulatedFieldDeclaration CreateEncapsulatedFieldMock(string targetID, string asTypeName, IVBComponent component, string modifiedPropertyName = null, bool encapsulateFlag = true )
        {
            var identifiers = new EncapsulationIdentifiers(targetID);
            var attributesMock = CreateAttributesMock(targetID, asTypeName, modifiedPropertyName);

            var mock = new Mock<IEncapsulatedFieldDeclaration>();
            mock.SetupGet(m => m.TargetID).Returns(identifiers.TargetFieldName);
            mock.SetupGet(m => m.NewFieldName).Returns(identifiers.Field);
            mock.SetupGet(m => m.PropertyName).Returns(modifiedPropertyName ?? identifiers.Property);
            mock.SetupGet(m => m.AsTypeName).Returns(asTypeName);
            mock.SetupGet(m => m.EncapsulateFlag).Returns(encapsulateFlag);
            mock.SetupGet(m => m.EncapsulationAttributes).Returns(attributesMock);
            mock.SetupGet(m => m.QualifiedModuleName).Returns(new QualifiedModuleName(component));
            return mock.Object;
        }

        private IFieldEncapsulationAttributes CreateAttributesMock(string targetID, string asTypeName, string modifiedPropertyName = null, bool encapsulateFlag = true)
        {
            var identifiers = new EncapsulationIdentifiers(targetID);
            var mock = new Mock<IFieldEncapsulationAttributes>();
            mock.SetupGet(m => m.TargetName).Returns(identifiers.TargetFieldName);
            mock.SetupGet(m => m.NewFieldName).Returns(identifiers.Field);
            mock.SetupGet(m => m.PropertyName).Returns(modifiedPropertyName ?? identifiers.Property);
            mock.SetupGet(m => m.AsTypeName).Returns(asTypeName);
            mock.SetupGet(m => m.EncapsulateFlag).Returns(encapsulateFlag);
            return mock.Object;
        }
    }
}
