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
    public sealed class ImplicitByRefParameterInspection : InspectionBase
    {
        public ImplicitByRefParameterInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var interfaceMembers = UserDeclarations.FindInterfaceImplementationMembers();
            var builtinEventHandlers = State.DeclarationFinder.FindEventHandlers();

            var issues = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>()
                .Where(item => item.IsImplicitByRef && !item.IsParamArray
                    && !IsIgnoringInspectionResultFor(item, AnnotationName)
                    && !interfaceMembers.Contains(item.ParentDeclaration)
                    && !builtinEventHandlers.Contains(item.ParentDeclaration))
                .ToList();

            return issues.Select(issue => new DeclarationInspectionResult(this, string.Format(InspectionsUI.ImplicitByRefParameterInspectionResultFormat, issue.IdentifierName), issue));
        }
    }
}
