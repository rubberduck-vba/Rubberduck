using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public class AttributeAnnotationProvider : IAttributeAnnotationProvider
    {
        public (AnnotationType annotationType, IReadOnlyList<string> values) ModuleAttributeAnnotation(
            string attributeName, 
            IReadOnlyList<string> attributeValues)
        {
            var moduleAnnotations = ModuleAnnotations();
            return AttributeAnnotation(
                moduleAnnotations, 
                attributeName, 
                attributeValues,
                AnnotationType.ModuleAttribute);
        }

        private (AnnotationType annotationType, IReadOnlyList<string> values) AttributeAnnotation(
            IReadOnlyList<AnnotationType> annotationTypes,
            string attributeName,
            IReadOnlyList<string> attributeValues,
            AnnotationType fallbackFlexibleAttributeAnnotationType)
        {
            var fixedValueAttributeAnnotation = FirstMatchingFixedAttributeValueAnnotation(annotationTypes, attributeName, attributeValues);
            if (fixedValueAttributeAnnotation != default)
            {
                return (fixedValueAttributeAnnotation, new List<string>());
            }

            var flexibleValueAttributeAnnotation = FirstMatchingFlexibleAttributeValueAnnotation(annotationTypes, attributeName, attributeValues.Count);
            if (flexibleValueAttributeAnnotation != default)
            {
                return (flexibleValueAttributeAnnotation, attributeValues);
            }

            var annotationValues = WithNewValuePrepended(attributeValues, attributeName);
            return (fallbackFlexibleAttributeAnnotationType, annotationValues);
        }

        private static IReadOnlyList<AnnotationType> ModuleAnnotations()
        {
            var type = typeof(AnnotationType);
            return Enum.GetValues(type)
                .Cast<AnnotationType>()
                .Where(annotationType => annotationType.HasFlag(AnnotationType.ModuleAnnotation))
                .ToList();
        }

        private static AnnotationType FirstMatchingFixedAttributeValueAnnotation(
            IEnumerable<AnnotationType> annotationTypes,
            string attributeName,
            IEnumerable<string> attributeValues)
        {
            var type = typeof(AnnotationType);
            return annotationTypes.FirstOrDefault(annotationType => type.GetField(Enum.GetName(type, annotationType))
                .GetCustomAttributes(false)
                .OfType<FixedAttributeValueAnnotationAttribute>()
                .Any(attribute => attribute.AttributeName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)
                                  && attribute.AttributeValues.SequenceEqual(attributeValues)));
        }

        private static AnnotationType FirstMatchingFlexibleAttributeValueAnnotation(
            IEnumerable<AnnotationType> annotationTypes,
            string attributeName,
            int valueCount)
        {
            var type = typeof(AnnotationType);
            return annotationTypes.FirstOrDefault(annotationType => type.GetField(Enum.GetName(type, annotationType))
                .GetCustomAttributes(false)
                .OfType<FlexibleAttributeValueAnnotationAttribute>()
                .Any(attribute => attribute.AttributeName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)
                                  && attribute.NumberOfParameters == valueCount));
        }

        private IReadOnlyList<string> WithNewValuePrepended(IReadOnlyList<string> oldList, string newValue)
        {
            var newList = oldList.ToList();
            newList.Insert(0, newValue);
            return newList;
        }

        public (AnnotationType annotationType, IReadOnlyList<string> values) MemberAttributeAnnotation(string attributeBaseName, IReadOnlyList<string> attributeValues)
        {
            var nonModuleAnnotations = NonModuleAnnotations();
            return AttributeAnnotation(
                nonModuleAnnotations,
                attributeBaseName,
                attributeValues,
                AnnotationType.MemberAttribute);
        }

        private static IReadOnlyList<AnnotationType> NonModuleAnnotations()
        {
            var type = typeof(AnnotationType);
            return Enum.GetValues(type)
                .Cast<AnnotationType>()
                .Where(annotationType => !annotationType.HasFlag(AnnotationType.ModuleAnnotation))
                .ToList();
        }
    }
}