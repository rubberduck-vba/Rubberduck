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

namespace Rubberduck.Inspections.Concrete
{
    public sealed class IllegalAnnotationInspection : InspectionBase
    {
        public IllegalAnnotationInspection(RubberduckParserState state)
            : base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var illegalAnnotations = new List<IAnnotation>();

            var userDeclarations = State.DeclarationFinder.AllUserDeclarations.ToList();
            var annotations = State.AllAnnotations;

            illegalAnnotations.AddRange(UnboundAnnotations(annotations, userDeclarations));
            illegalAnnotations.AddRange(NonModuleAnnotationsOnModules(userDeclarations));
            illegalAnnotations.AddRange(NonMemberAnnotationsOnMembers(userDeclarations));

            return illegalAnnotations.Select(annotation => 
                new QualifiedContextInspectionResult(
                    this, 
                    string.Format(InspectionResults.IllegalAnnotationInspection, annotation.Context.annotationName().GetText()), 
                    new QualifiedContext(annotation.QualifiedSelection.QualifiedName, annotation.Context)));
        }

        private ICollection<IAnnotation> UnboundAnnotations(IEnumerable<IAnnotation> annotations, IEnumerable<Declaration> userDeclarations)
        {
            var boundAnnotations = userDeclarations.SelectMany(declaration => declaration.Annotations)
                .ToDictionary(annotation => annotation.QualifiedSelection);
            
            return annotations.Where(annotation => !boundAnnotations.ContainsKey(annotation.QualifiedSelection)).ToList();
        }

        private ICollection<IAnnotation> NonModuleAnnotationsOnModules(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module))
                .SelectMany(moduleDeclaration => moduleDeclaration.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.ModuleAnnotation))
                .ToList();
        }

        private ICollection<IAnnotation> NonMemberAnnotationsOnMembers(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => !declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.DeclarationType != DeclarationType.Project)
                .SelectMany(member => member.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.MemberAnnotation))
                .ToList();
        }
    }
}