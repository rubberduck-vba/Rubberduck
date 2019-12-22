using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using RubberduckTests.Mocks;

namespace RubberduckTests.Annotations
{
    [TestFixture]
    [Category("Annotations")]
    public class AttributeAnnotationProviderTests
    {
        [Test]
        public void FindMemberAnnotationForRandomAttributeReturnsMemberAttributeAnnotation()
        {
            var attributeName = "VB_Whatever";
            var attributeValues = new List<string>{ "\"SomeValue\"" };

            var expectedAnnotationType = "MemberAttribute";
            var expectedValues = new List<string>{ "VB_Whatever", "\"SomeValue\"" };

            var attributeAnnotationProvider = GetAnnotationProvider();
            var (actualAnnotationInfo, actualValues) = attributeAnnotationProvider.MemberAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, actualAnnotationInfo.Name);
            AssertEqual(expectedValues, actualValues);
        }

        [Test]
        public void FindModuleAnnotationForRandomAttributeReturnsModuleAttributeAnnotation()
        {
            var attributeName = "VB_Whatever";
            var attributeValues = new List<string> { "\"SomeValue\"" };

            var expectedAnnotationType = "ModuleAttribute";
            var expectedValues = new List<string> { "VB_Whatever", "\"SomeValue\"" };

            var attributeAnnotationProvider = GetAnnotationProvider();
            var (annotationInfo, actualValues) = attributeAnnotationProvider.ModuleAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, annotationInfo.Name);
            AssertEqual(expectedValues, actualValues);
        }

        [TestCase("VB_Description", "\"SomeDescription\"", "ModuleDescription", "\"SomeDescription\"")]
        [TestCase("VB_Exposed", "True", "Exposed")]
        [TestCase("VB_PredeclaredId", "True", "PredeclaredId")]
        public void ModuleAttributeAnnotationReturnsSpecializedAnnotationsWhereApplicable(string attributeName, string annotationValue, string expectedAnnotationType, string expectedValue = null)
        {
            var attributeValues = new List<string> { annotationValue };
            var expectedValues = expectedValue != null
                                    ? new List<string> { expectedValue }
                                    : new List<string>();
            
            var attributeAnnotationProvider = GetAnnotationProvider();
            var (annotationInfo, actualValues) = attributeAnnotationProvider.ModuleAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, annotationInfo.Name);
            AssertEqual(expectedValues, actualValues);
        }

        [TestCase("VB_ProcData.VB_Invoke_Func", @"A\n14", "ExcelHotkey", "A")]
        [TestCase("VB_Description", "\"SomeDescription\"", "Description", "\"SomeDescription\"")]
        [TestCase("VB_VarDescription", "\"SomeDescription\"", "VariableDescription", "\"SomeDescription\"")]
        [TestCase("VB_UserMemId", "0", "DefaultMember")]
        [TestCase("VB_UserMemId", "-4", "Enumerator")]
        public void MemberAttributeAnnotationReturnsSpecializedAnnotationsWhereApplicable(string attributeName, string attributeValue, string expectedAnnotationType, string expectedValue = null)
        {
            var attributeValues = new List<string> { attributeValue };
            var expectedValues = expectedValue != null
                ? new List<string> { expectedValue }
                : new List<string>();

            var attributeAnnotationProvider = GetAnnotationProvider();
            var (annotationInfo, actualValues) = attributeAnnotationProvider.MemberAttributeAnnotation(attributeName, attributeValues);

            Assert.AreEqual(expectedAnnotationType, annotationInfo.Name);
            AssertEqual(expectedValues, actualValues);
        }

        private AttributeAnnotationProvider GetAnnotationProvider()
        {
            return new AttributeAnnotationProvider(MockParser.WellKnownAnnotations().OfType<IAttributeAnnotation>());
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