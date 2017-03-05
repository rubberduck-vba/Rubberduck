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
    public sealed class ConstantNotUsedInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;

        public ConstantNotUsedInspection(RubberduckParserState state, IMessageBox messageBox)
            : base(state)
        {
            _messageBox = messageBox;
        }

        public override string Meta { get { return InspectionsUI.ConstantNotUsedInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ConstantNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var results = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Constant)
                .Where(declaration => !declaration.References.Any()
                    && !IsIgnoringInspectionResultFor(declaration, AnnotationName))
                .ToList();

            return results.Select(issue => 
                new IdentifierNotUsedInspectionResult(this, issue, ((dynamic)issue.Context).identifier(), issue.QualifiedName.QualifiedModuleName));
        }
    }
}
