using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ParameterNotUsedInspection : InspectionBase
    {
        public ParameterNotUsedInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers();
            var interfaceImplementationMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers();

            var handlers = State.DeclarationFinder.FindEventHandlers();

            var parameters = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>()
                .Where(parameter => !parameter.References.Any() && !IsIgnoringInspectionResultFor(parameter, AnnotationName)
                                    && parameter.ParentDeclaration.DeclarationType != DeclarationType.Event
                                    && parameter.ParentDeclaration.DeclarationType != DeclarationType.LibraryFunction
                                    && parameter.ParentDeclaration.DeclarationType != DeclarationType.LibraryProcedure
                                    && !interfaceMembers.Contains(parameter.ParentDeclaration)
                                    && !handlers.Contains(parameter.ParentDeclaration))
                .ToList();

            var issues = from issue in parameters
                let isInterfaceImplementationMember = interfaceImplementationMembers.Contains(issue.ParentDeclaration)
                select new DeclarationInspectionResult(this, string.Format(InspectionsUI.ParameterNotUsedInspectionResultFormat, issue.IdentifierName).Capitalize(), issue);

            return issues;
        }
    }
}
