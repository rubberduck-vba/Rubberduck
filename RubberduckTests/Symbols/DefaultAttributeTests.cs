using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class DefaultAttributeTests
    {
        [Test]
        public void NoAttributeIsConsideredADefaultModuleIfTheModuleIsNeitherModuleClassNorUserForm()
        {
            var componentType = ComponentType.Document;
            var attribute = "VB_Name";
            var attributeValues = new List<string> { "\"Name\"" };

            var isDefaultAttribute = Attributes.IsDefaultAttribute(componentType, attribute, attributeValues);

            Assert.IsFalse(isDefaultAttribute);
        }

        [TestCase(ComponentType.StandardModule)]
        [TestCase(ComponentType.ClassModule)]
        [TestCase(ComponentType.UserForm)]
        public void NameIsConsideredADefaultAttributeIrrespectiveOdValue(ComponentType componentType)
        {
            var attribute = "VB_Name";
            var attributeValues = new List<string> { "\"Whatever\"" };

            var isDefaultAttribute = Attributes.IsDefaultAttribute(componentType, attribute, attributeValues);

            Assert.IsTrue(isDefaultAttribute);
        }

        [TestCase(ComponentType.StandardModule, "True", false)]
        [TestCase(ComponentType.StandardModule, "False", false)]
        [TestCase(ComponentType.ClassModule, "True", false)]
        [TestCase(ComponentType.ClassModule, "False", true)]
        [TestCase(ComponentType.UserForm, "True", false)]
        [TestCase(ComponentType.UserForm, "False", true)]
        public void ExposedIsConsideredADefaultAttributeForTheRightValue(ComponentType componentType, string attributeValue, bool isDefault)
        {
            var attribute = "VB_Exposed";
            var attributeValues = new List<string> { attributeValue };

            var isDefaultAttribute = Attributes.IsDefaultAttribute(componentType, attribute, attributeValues);

            Assert.AreEqual(isDefault, isDefaultAttribute);
        }

        [TestCase(ComponentType.StandardModule, "True", false)]
        [TestCase(ComponentType.StandardModule, "False", false)]
        [TestCase(ComponentType.ClassModule, "True", false)]
        [TestCase(ComponentType.ClassModule, "False", true)]
        [TestCase(ComponentType.UserForm, "True", false)]
        [TestCase(ComponentType.UserForm, "False", true)]
        public void CreateableIsConsideredADefaultAttributeForTheRightValue(ComponentType componentType, string attributeValue, bool isDefault)
        {
            var attribute = "VB_Creatable";
            var attributeValues = new List<string> { attributeValue };

            var isDefaultAttribute = Attributes.IsDefaultAttribute(componentType, attribute, attributeValues);

            Assert.AreEqual(isDefault, isDefaultAttribute);
        }

        [TestCase(ComponentType.StandardModule, "True", false)]
        [TestCase(ComponentType.StandardModule, "False", false)]
        [TestCase(ComponentType.ClassModule, "True", false)]
        [TestCase(ComponentType.ClassModule, "False", true)]
        [TestCase(ComponentType.UserForm, "True", false)]
        [TestCase(ComponentType.UserForm, "False", true)]
        public void GlobalNameSpaceIsConsideredADefaultAttributeForTheRightValue(ComponentType componentType, string attributeValue, bool isDefault)
        {
            var attribute = "VB_GlobalNameSpace";
            var attributeValues = new List<string> { attributeValue };

            var isDefaultAttribute = Attributes.IsDefaultAttribute(componentType, attribute, attributeValues);

            Assert.AreEqual(isDefault, isDefaultAttribute);
        }

        [TestCase(ComponentType.StandardModule, "True", false)]
        [TestCase(ComponentType.StandardModule, "False", false)]
        [TestCase(ComponentType.ClassModule, "True", false)]
        [TestCase(ComponentType.ClassModule, "False", true)]
        [TestCase(ComponentType.UserForm, "True", true)]
        [TestCase(ComponentType.UserForm, "False", false)]
        public void PredeclaredIdIsConsideredADefaultAttributeForTheRightValue(ComponentType componentType, string attributeValue, bool isDefault)
        {
            var attribute = "VB_PredeclaredId";
            var attributeValues = new List<string> { attributeValue };

            var isDefaultAttribute = Attributes.IsDefaultAttribute(componentType, attribute, attributeValues);

            Assert.AreEqual(isDefault, isDefaultAttribute);
        }
    }
}