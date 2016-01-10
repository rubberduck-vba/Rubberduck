using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class MultipleNamespaceDeclarationInspection : InspectionBase
    {
        public MultipleNamespaceDeclarationInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return InspectionsUI.MultipleNamespaceDeclarationInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override bool CanIgnoreOnce { get { return false; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations.Where(declaration =>
                 (declaration.DeclarationType == DeclarationType.Class
                || declaration.DeclarationType == DeclarationType.Module)
                && declaration.Annotations.Split('\n').Count(annotation =>
                    annotation.StartsWith(Parsing.Grammar.Annotations.AnnotationMarker +
                                          Parsing.Grammar.Annotations.Namespace)) > 1);
            return issues.Select(issue =>
                new MultipleNamespaceDeclarationInspectionResult(this, string.Format(Description, issue.ComponentName), issue));
        }
    }

    public class MultipleNamespaceDeclarationInspectionResult : CodeInspectionResultBase
    {
        public MultipleNamespaceDeclarationInspectionResult(IInspection inspection, string result, Declaration target) 
            : base(inspection, result, target)
        {
        }
    }
}