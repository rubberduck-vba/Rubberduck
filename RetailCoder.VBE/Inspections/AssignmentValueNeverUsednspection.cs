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
            var issues = UserDeclarations.Where(declaration =>
                        declaration.DeclarationType == DeclarationType.Variable ||
                        declaration.DeclarationType == DeclarationType.Parameter).ToList();

            var unusedAssignments = new List<IdentifierReference>();
            
            foreach (var referencesByScope in issues.Select(issue => issue.References
                                                                   .GroupBy(ParentDeclaration)
                                                                   .Select(g => g.OrderBy(o => o.Selection.StartLine).ThenBy(t => t.Selection.StartColumn))
                                                                   .ToList()))
            {
                foreach (var references in referencesByScope.Select(r => r.ToList()))
                {
                    var lastNonAssignmentReference = references.FindLastIndex(r => !r.IsAssignment);
                    for (var i = 0; i < references.Count - 1; i++)
                    {
                        var currentReference = references[i];
                        var nextReference = references[i + 1];

                        if (currentReference.IsAssignment && nextReference.IsAssignment &&
                            !unusedAssignments.Contains(currentReference))
                        {
                            unusedAssignments.Add(currentReference);
                        }

                        var isLastAssignmentToFieldInScope = new[] {DeclarationType.Class, DeclarationType.Module}
                                                                .Contains(currentReference.Declaration.ParentDeclaration.DeclarationType) &&
                                                            i + 1 == references.Count - 1;  // here, we are checking the next reference

                        if (!isLastAssignmentToFieldInScope &&
                            i + 1 > lastNonAssignmentReference &&
                            !unusedAssignments.Contains(nextReference))
                        {
                            unusedAssignments.Add(nextReference);
                        }
                    }
                }
            }

            return unusedAssignments.Select(r => new AssignmentValueNeverUsedInspectionResult(this, r));
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
