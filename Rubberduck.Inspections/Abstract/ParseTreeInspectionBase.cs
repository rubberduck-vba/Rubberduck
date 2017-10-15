using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class ParseTreeInspectionBase : InspectionBase, IParseTreeInspection
    {
        protected ParseTreeInspectionBase(RubberduckParserState state, CodeInspectionSeverity severity = CodeInspectionSeverity.Warning)
            : base(state, severity) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;
        public abstract IInspectionListener Listener { get; }
        public virtual ParsePass Pass => ParsePass.CodePanePass;
    }
}