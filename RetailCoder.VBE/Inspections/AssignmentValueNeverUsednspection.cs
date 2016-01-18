using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class AssignmentValueNeverUsedInspection : InspectionBase
    {
        public AssignmentValueNeverUsedInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return InspectionsUI.AssignmentValueNeverUsedInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues =
                UserDeclarations.Where(
                    declaration =>
                        declaration.DeclarationType == DeclarationType.Variable ||
                        declaration.DeclarationType == DeclarationType.Parameter).ToList();

            var unassignedReferences = new List<IdentifierReference>();

            foreach (var references in issues.Select(issue => issue.References
                                                                   .OrderBy(o => ParentDeclaration(o).Scope)
                                                                   .ThenBy(t => t.Selection.StartLine)
                                                                   .ThenBy(t => t.Selection.StartColumn)
                                                                   .ToList()))
            {
                var lastNonAssignmentReference = references.FindLastIndex(r => !r.IsAssignment);
                for (var i = 1; i < references.Count; i++)
                {
                    if (references[i].IsAssignment && references[i - 1].IsAssignment &&
                        ParentDeclaration(references[i]).Scope == ParentDeclaration(references[i - 1]).Scope &&
                        !unassignedReferences.Contains(references[i - 1]))
                    {
                        unassignedReferences.Add(references[i - 1]);
                    }

                    var isLastReferenceToFieldInScope = new[] {DeclarationType.Class, DeclarationType.Module}.Contains(
                        references[i].Declaration.ParentDeclaration.DeclarationType) &&
                              (i == references.Count - 1 || references[i].ParentScope != references[i + 1].ParentScope);

                    if (!isLastReferenceToFieldInScope &&
                        i > lastNonAssignmentReference &&
                        !unassignedReferences.Contains(references[i]))
                    {
                        unassignedReferences.Add(references[i]);
                    }
                }
            }

            return unassignedReferences.Select(r => new AssignmentValueNeverUsedInspectionResult(this, r));
        }

        private Declaration ParentDeclaration(IdentifierReference reference)
        {
            var declarationTypes = new[] { DeclarationType.Function, DeclarationType.Procedure, DeclarationType.Property };

            return UserDeclarations.SingleOrDefault(d =>
                        d.Scope == reference.ParentScope && declarationTypes.Contains(d.DeclarationType) &&
                        d.Project == reference.QualifiedModuleName.Project);
        }
    }
}
