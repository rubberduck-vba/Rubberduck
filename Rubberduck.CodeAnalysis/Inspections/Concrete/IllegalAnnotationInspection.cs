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
            var illegalAnnotations = new List<IAnnotation>();

            var userDeclarations = State.DeclarationFinder.AllUserDeclarations.ToList();
            var identifierReferences = State.DeclarationFinder.AllIdentifierReferences().ToList();
            var annotations = State.AllAnnotations;

            illegalAnnotations.AddRange(UnboundAnnotations(annotations, userDeclarations, identifierReferences));
            illegalAnnotations.AddRange(NonIdentifierAnnotationsOnIdentifiers(identifierReferences));
            illegalAnnotations.AddRange(NonModuleAnnotationsOnModules(userDeclarations));
            illegalAnnotations.AddRange(NonMemberAnnotationsOnMembers(userDeclarations));
            illegalAnnotations.AddRange(NonVariableAnnotationsOnVariables(userDeclarations));
            illegalAnnotations.AddRange(NonGeneralAnnotationsWhereOnlyGeneralAnnotationsBelong(userDeclarations));

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

        private static ICollection<IAnnotation> NonIdentifierAnnotationsOnIdentifiers(IEnumerable<IdentifierReference> identifierReferences)
        {
            return identifierReferences
                .SelectMany(reference => reference.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.IdentifierAnnotation))
                .ToList();
        }

        private static ICollection<IAnnotation> NonModuleAnnotationsOnModules(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module))
                .SelectMany(moduleDeclaration => moduleDeclaration.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.ModuleAnnotation))
                .ToList();
        }

        private static ICollection<IAnnotation> NonMemberAnnotationsOnMembers(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Member))
                .SelectMany(member => member.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.MemberAnnotation))
                .ToList();
        }

        private static ICollection<IAnnotation> NonVariableAnnotationsOnVariables(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => VariableAnnotationDeclarationTypes.Contains(declaration.DeclarationType))
                .SelectMany(declaration => declaration.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.VariableAnnotation))
                .ToList();
        }

        private static readonly HashSet<DeclarationType> VariableAnnotationDeclarationTypes = new HashSet<DeclarationType>()
        {
            DeclarationType.Variable,
            DeclarationType.Control,
            DeclarationType.Constant,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember
        };

        private static ICollection<IAnnotation> NonGeneralAnnotationsWhereOnlyGeneralAnnotationsBelong(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => !declaration.DeclarationType.HasFlag(DeclarationType.Module) 
                                      && !declaration.DeclarationType.HasFlag(DeclarationType.Member) 
                                      && !VariableAnnotationDeclarationTypes.Contains(declaration.DeclarationType) 
                                      && declaration.DeclarationType != DeclarationType.Project)
                .SelectMany(member => member.Annotations)
                .Where(annotation => !annotation.AnnotationType.HasFlag(AnnotationType.GeneralAnnotation))
                .ToList();
        }
    }
}