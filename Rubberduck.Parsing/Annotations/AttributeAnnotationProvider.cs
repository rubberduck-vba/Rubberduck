using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Rubberduck.Parsing.Annotations.Concrete;

namespace Rubberduck.Parsing.Annotations
{
    public class AttributeAnnotationProvider : IAttributeAnnotationProvider
    {
        // I want to const this, but can't
        private readonly AnnotationTarget [] distinctTargets = new AnnotationTarget[] { AnnotationTarget.Identifier, AnnotationTarget.Member, AnnotationTarget.Module, AnnotationTarget.Variable };
        private readonly Dictionary<AnnotationTarget, List<IAttributeAnnotation>> annotationInfoByTarget
            = new Dictionary<AnnotationTarget, List<IAttributeAnnotation>>();

        private readonly IAttributeAnnotation memberFallback = new MemberAttributeAnnotation();
        private readonly IAttributeAnnotation moduleFallback = new ModuleAttributeAnnotation();
        
        public AttributeAnnotationProvider(IEnumerable<IAttributeAnnotation> attributeAnnotations)
        {
            // set up empty lists to put information into
            foreach (var validTarget in distinctTargets)
            {
                annotationInfoByTarget[validTarget] = new List<IAttributeAnnotation>();
            }

            foreach (var annotation in attributeAnnotations)
            {
                foreach (var validTarget in distinctTargets)
                {
                    if (annotation.Target.HasFlag(validTarget))
                    {
                        annotationInfoByTarget[validTarget].Add(annotation);
                    }
                }
            }
        }

        public (IAttributeAnnotation annotation, IReadOnlyList<string> annotationValues) MemberAttributeAnnotation(string attributeBaseName, IReadOnlyList<string> attributeValues)
        {
            // go through all non-module annotations (contrary to only member annotations)
            var memberAnnotationTypes = annotationInfoByTarget[AnnotationTarget.Member]
                .Concat(annotationInfoByTarget[AnnotationTarget.Variable])
                .Concat(annotationInfoByTarget[AnnotationTarget.Identifier]);
            foreach (var annotation in memberAnnotationTypes)
            {
                if (annotation.MatchesAttributeDefinition(attributeBaseName, attributeValues))
                {
                    return (annotation, annotation.AttributeToAnnotationValues(attributeValues));
                }
            }
            var fallbackAttributeArguments = new[] { attributeBaseName }.Concat(attributeValues);
            return (memberFallback, memberFallback.AttributeToAnnotationValues(fallbackAttributeArguments.ToList()));
        }

        public (IAttributeAnnotation annotation, IReadOnlyList<string> annotationValues) ModuleAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues)
        {
            var moduleAnnotationTypes = annotationInfoByTarget[AnnotationTarget.Module];
            foreach (var annotation in moduleAnnotationTypes)
            {
                if (annotation.MatchesAttributeDefinition(attributeName, attributeValues))
                {
                    return (annotation, annotation.AttributeToAnnotationValues(attributeValues));
                }
            }
            var fallbackAttributeArguments = new[] { attributeName }.Concat(attributeValues);
            return (moduleFallback, moduleFallback.AttributeToAnnotationValues(fallbackAttributeArguments.ToList()));
        }
    }
}