using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class ParameterNotUsedInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;

        public ParameterNotUsedInspection(RubberduckParserState state, IMessageBox messageBox)
            : base(state)
        {
            _messageBox = messageBox;
        }

        public override string Meta { get { return InspectionsUI.ParameterNotUsedInspectionName; }}
        public override string Description { get { return InspectionsUI.ParameterNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
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
                select new ParameterNotUsedInspectionResult(this, issue, isInterfaceImplementationMember, issue.Project.VBE, State, _messageBox);

            return issues.ToList();
        }
    }
}
