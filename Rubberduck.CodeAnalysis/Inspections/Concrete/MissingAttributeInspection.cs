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
            foreach (var declaration in declarationsWithAttributeAnnotations)
            {
                foreach(var annotation in declaration.Annotations.Where(annotation => annotation.AnnotationType.HasFlag(AnnotationType.Attribute)))
                {
                    if (MissesCorrespondingAttribute(declaration, annotation))
                    {
                        var description = string.Format(InspectionResults.MissingAttributeInspection, declaration.IdentifierName,
                            annotation.AnnotationType.ToString());
                        results.Add(new DeclarationInspectionResult(this, description, declaration, new QualifiedContext(declaration.QualifiedModuleName, annotation.Context), annotation));
                    }
                }
            }

            return results;
        }

        private static bool MissesCorrespondingAttribute(Declaration declaration, IAnnotation annotation)
        {
            return !declaration.Attributes.HasAttributeFor(annotation.AnnotationType);
        }
    }
}