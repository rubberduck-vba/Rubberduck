using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Rubberduck.Parsing.Annotations
{
    public class AttributeAnnotationProvider : IAttributeAnnotationProvider
    {
        // I want to const this, but can't
        private readonly AnnotationTarget [] distinctTargets = new AnnotationTarget[] { AnnotationTarget.Identifier, AnnotationTarget.Member, AnnotationTarget.Module, AnnotationTarget.Variable };
        private readonly Dictionary<AnnotationTarget, List<Type>> annotationInfoByTarget
            = new Dictionary<AnnotationTarget, List<Type>>();

        // FIXME make sure only AttributeAnnotations are injected here
        public AttributeAnnotationProvider(IEnumerable<Type> attributeAnnotationTypes)
        {
            // set up empty lists to put information into
            foreach (var validTarget in distinctTargets)
            {
                annotationInfoByTarget[validTarget] = new List<Type>();
            }
            // we're defensively filtering, but theoretically this might be CW's job?
            foreach (var annotationType in attributeAnnotationTypes.Where(type => type.GetInterfaces().Contains(typeof(IAttributeAnnotation))))
            {
                // Extract the static information about the annotation type from it's AnnotationAttribute
                var staticInfo = annotationType.GetCustomAttributes(false)
                    .OfType<AnnotationAttribute>()
                    .Single();
                foreach (var validTarget in distinctTargets)
                {
                    if (staticInfo.Target.HasFlag(validTarget))
                    {
                        annotationInfoByTarget[validTarget].Add(annotationType);
                    }
                }
            }
        }

        public (AnnotationAttribute annotationInfo, IReadOnlyList<string> values) MemberAttributeAnnotation(string attributeBaseName, IReadOnlyList<string> attributeValues)
        {
            // quasi-const
            var fallbackType = typeof(MemberAttributeAnnotation);
            // go through all non-module annotations (contrary to only member annotations)
            var memberAnnotationTypes = annotationInfoByTarget[AnnotationTarget.Member]
                .Concat(annotationInfoByTarget[AnnotationTarget.Variable])
                .Concat(annotationInfoByTarget[AnnotationTarget.Identifier]);
            foreach (var type in memberAnnotationTypes)
            {
                if (MatchesAttributeNameAndValue(type, attributeBaseName, attributeValues, out var codePassAnnotationValues))
                {
                    return (GetAttribute<AnnotationAttribute>(type), codePassAnnotationValues);
                }
            }
            return BuildFallback(attributeBaseName, attributeValues, fallbackType);
        }

        public (AnnotationAttribute annotationInfo, IReadOnlyList<string> values) ModuleAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues)
        {
            // quasi-const
            var fallbackType = typeof(ModuleAttributeAnnotation);
            var moduleAnnotationTypes = annotationInfoByTarget[AnnotationTarget.Module];
            foreach (var type in moduleAnnotationTypes)
            {
                if (MatchesAttributeNameAndValue(type, attributeName, attributeValues, out var codePassAnnotationValues))
                {
                    return (GetAttribute<AnnotationAttribute>(type), codePassAnnotationValues);
                }
            }
            return BuildFallback(attributeName, attributeValues, fallbackType);
        }

        private bool MatchesAttributeNameAndValue(Type type, string attributeName, IReadOnlyList<string> attributeValues, out IReadOnlyList<string> codePassAnnotationValues)
        {
            codePassAnnotationValues = attributeValues;
            if (typeof(FlexibleAttributeAnnotationBase).IsAssignableFrom(type))
            {
                // this is always the fallback case, which must only be accepted if all other options are exhausted.
                return false;
            }
            if (typeof(FixedAttributeValueAnnotationBase).IsAssignableFrom(type))
            {
                var attributeInfo = GetAttribute<FixedAttributeValueAnnotationAttribute>(type);
                if (attributeInfo.AttributeName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)
                    && attributeInfo.AttributeValues.SequenceEqual(attributeValues))
                {
                    // there is no way to set a value in the annotation, therefore we discard the attribute values
                    codePassAnnotationValues = new List<string>();
                    return true;
                }
            }
            if (typeof(FlexibleAttributeValueAnnotationBase).IsAssignableFrom(type))
            {
                // obtain flexible attribute information
                var attributeInfo = GetAttribute<FlexibleAttributeValueAnnotationAttribute>(type);
                if (attributeInfo.AttributeName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)
                    && attributeInfo.NumberOfParameters == attributeValues.Count)
                {
                    if (attributeInfo.HasCustomTransformation)
                    {
                        try {
                        // dispatch to custom transformation
                            codePassAnnotationValues = ((IEnumerable<string>)type.GetMethod("TransformToAnnotationValues", new[] { typeof(IEnumerable<string>) })
                                .Invoke(null, new[] { attributeValues })).ToList();
                        }
                        catch (Exception)
                        {
                            codePassAnnotationValues = attributeValues;
                        }
                    }
                    return true;
                }
            }
            return false;
        }

        private (AnnotationAttribute annotationInfo, IReadOnlyList<string> values) BuildFallback(string attributeBaseName, IReadOnlyList<string> attributeValues, Type fallbackType)
        {
            var fallbackValues = new[] { attributeBaseName }.Concat(attributeValues).ToList();
            return (GetAttribute<AnnotationAttribute>(fallbackType), fallbackValues);
        }

        private static T GetAttribute<T>(Type annotationType)
        {
            return annotationType.GetCustomAttributes(false)
                .OfType<T>()
                .Single();
        }
    }
}