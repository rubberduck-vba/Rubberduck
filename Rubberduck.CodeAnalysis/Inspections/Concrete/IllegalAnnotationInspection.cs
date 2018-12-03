using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class IllegalAnnotationInspection : InspectionBase
    {
        public IllegalAnnotationInspection(RubberduckParserState state)
            : base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var userDeclarations = State.DeclarationFinder.AllUserDeclarations.ToList();
            var identifierReferences = State.DeclarationFinder.AllIdentifierReferences().ToList();
            var annotations = State.AllAnnotations;

            var illegalAnnotations = UnboundAnnotations(annotations, userDeclarations, identifierReferences);

            return illegalAnnotations.Select(annotation => 
                new QualifiedContextInspectionResult(
                    this, 
                    string.Format(InspectionResults.IllegalAnnotationInspection, annotation.Context.annotationName().GetText()), 
                    new QualifiedContext(annotation.QualifiedSelection.QualifiedName, annotation.Context)));
        }

        private static ICollection<IAnnotation> UnboundAnnotations(IEnumerable<IAnnotation> annotations, IEnumerable<Declaration> userDeclarations, IEnumerable<IdentifierReference> identifierReferences)
        {
            var boundAnnotationsSelections = userDeclarations
                .SelectMany(declaration => declaration.Annotations)
                .Concat(identifierReferences.SelectMany(reference => reference.Annotations))
                .Select(annotation => annotation.QualifiedSelection)
                .ToHashSet();
            
            return annotations.Where(annotation => !boundAnnotationsSelections.Contains(annotation.QualifiedSelection)).ToList();
        }
    }
}