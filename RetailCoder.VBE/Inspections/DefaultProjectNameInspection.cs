using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class DefaultProjectNameInspection : InspectionBase
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public DefaultProjectNameInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _wrapperFactory = new CodePaneWrapperFactory();
        }

        public override string Meta { get { return InspectionsUI.DefaultProjectNameInspectionMeta; } }
        public override string Description { get { return InspectionsUI.DefaultProjectNameInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations
                            .Where(declaration => declaration.DeclarationType == DeclarationType.Project
                                                && declaration.IdentifierName.StartsWith("VBAProject"))
                            .Select(issue => new DefaultProjectNameInspectionResult(this, issue, State, _wrapperFactory))
                            .ToList();

            return issues;
        }
    }
}
