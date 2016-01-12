using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class DefaultProjectNameInspection : InspectionBase
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public DefaultProjectNameInspection(RubberduckParserState state)
            : base(state)
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public override string Description { get { return RubberduckUI.GenericProjectName_; } }
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
