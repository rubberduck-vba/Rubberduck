using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    [CannotAnnotate]
    public sealed class MissingAttributeInspection : InspectionBase
    {
        public MissingAttributeInspection(RubberduckParserState state)
            : base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarationsWithAttributeAnnotations = State.DeclarationFinder.AllUserDeclarations
                .Where(declaration => declaration.Annotations.Any(annotation => annotation.AnnotationType.HasFlag(AnnotationType.Attribute)));
            var results = new List<DeclarationInspectionResult>();
            foreach (var declaration in declarationsWithAttributeAnnotations.Where(decl => decl.QualifiedModuleName.ComponentType != ComponentType.Document))
            {
                foreach(var annotation in declaration.Annotations.OfType<IAttributeAnnotation>())
                {
                    if (MissesCorrespondingAttribute(declaration, annotation))
                    {
                        var description = string.Format(InspectionResults.MissingAttributeInspection, declaration.IdentifierName,
                            annotation.AnnotationType.ToString());

                        var result = new DeclarationInspectionResult(this, description, declaration,
                            new QualifiedContext(declaration.QualifiedModuleName, annotation.Context));
                        result.Properties.Annotation = annotation;

                        results.Add(result);
                    }
                }
            }

            return results;
        }

        private static bool MissesCorrespondingAttribute(Declaration declaration, IAttributeAnnotation annotation)
        {
            if (string.IsNullOrEmpty(annotation.Attribute))
            {
                return false;
            }
            return declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? !declaration.Attributes.HasAttributeFor(annotation)
                : !declaration.Attributes.HasAttributeFor(annotation, declaration.IdentifierName);
        }
    }
}