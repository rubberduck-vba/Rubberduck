using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitByRefParameterInspection : InspectionBase
    {
        public ImplicitByRefParameterInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.ImplicitByRefParameterInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitByRefParameterInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
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

            return issues.Select(issue => new ImplicitByRefParameterInspectionResult(this, issue));
        }
    }
}
