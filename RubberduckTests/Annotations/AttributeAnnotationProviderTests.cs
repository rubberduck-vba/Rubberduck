using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;

namespace RubberduckTests.Annotations
{
    [TestFixture]
    public class AttributeAnnotationProviderTests
    {
        [Test]
        public void FindMemberAnnotationForRandomAttributeReturnsMemberAttributeAnnotation()
        {
            var attributeName = "VB_Whatever";
            var attributeValues = new List<string>{"SomeValue"};

            var expectedAnnotationType = AnnotationType.MemberAttribute;
            var expectedValues = new List<string>{"VB_Whatever", "SomeValue"};

            var attributeAnnotationProvider = new AttributeAnnotationProvider();
            var (actualAnnotationType, actualValues) = attributeAnnotationProvider.MemberAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, actualAnnotationType);
            AssertEqual(expectedValues, actualValues);
        }

        [Test]
        public void FindModuleAnnotationForRandomAttributeReturnsModuleAttributeAnnotation()
        {
            var attributeName = "VB_Whatever";
            var attributeValues = new List<string> { "SomeValue" };

            var expectedAnnotationType = AnnotationType.ModuleAttribute;
            var expectedValues = new List<string> { "VB_Whatever", "SomeValue" };

            var attributeAnnotationProvider = new AttributeAnnotationProvider();
            var (actualAnnotationType, actualValues) = attributeAnnotationProvider.ModuleAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, actualAnnotationType);
            AssertEqual(expectedValues, actualValues);
        }

        [TestCase("VB_Description", "\"SomeDescription\"", AnnotationType.ModuleDescription, "\"SomeDescription\"")]
        [TestCase("VB_Exposed", "True", AnnotationType.Exposed)]
        [TestCase("VB_PredeclaredId", "True", AnnotationType.PredeclaredId)]
        public void ModuleAttributeAnnotationReturnsSpecializedAnnotationsWhereApplicable(string attributeName, string annotationValue, AnnotationType expectedAnnotationType, string expectedValue = null)
        {
            var attributeValues = new List<string> { annotationValue };
            var expectedValues = expectedValue != null
                                    ? new List<string> { expectedValue }
                                    : new List<string>();

            var attributeAnnotationProvider = new AttributeAnnotationProvider();
            var (actualAnnotationType, actualValues) = attributeAnnotationProvider.ModuleAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, actualAnnotationType);
            AssertEqual(expectedValues, actualValues);
        }

        [TestCase("VB_Description", "\"SomeDescription\"", AnnotationType.Description, "\"SomeDescription\"")]
        [TestCase("VB_VarDescription", "\"SomeDescription\"", AnnotationType.VariableDescription, "\"SomeDescription\"")]
        [TestCase("VB_UserMemId", "0", AnnotationType.DefaultMember)]
        [TestCase("VB_UserMemId", "-4", AnnotationType.Enumerator)]
        public void MemberAttributeAnnotationReturnsSpecializedAnnotationsWhereApplicable(string attributeName, string annotationValue, AnnotationType expectedAnnotationType, string expectedValue = null)
        {
            var attributeValues = new List<string> { annotationValue };
            var expectedValues = expectedValue != null
                ? new List<string> { expectedValue }
                : new List<string>();

            var attributeAnnotationProvider = new AttributeAnnotationProvider();
            var (actualAnnotationType, actualValues) = attributeAnnotationProvider.MemberAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, actualAnnotationType);
            AssertEqual(expectedValues, actualValues);
        }

        private static void AssertEqual(IReadOnlyList<string> expectedList, IReadOnlyList<string> actualList)
        {
            Assert.AreEqual(expectedList.Count, actualList.Count);
            for (int i = 0; i < expectedList.Count; i++)
            {
                Assert.AreEqual(expectedList[i], actualList[i]);
            }
        }
    }
}