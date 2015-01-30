using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class VariableTypeNotDeclaredInspection : IInspection
    {
        public VariableTypeNotDeclaredInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.VariableTypeNotDeclared; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IDictionary<QualifiedModuleName, IParseTree> nodes)
        {
            var targets = nodes.SelectMany(kvp => kvp.Value.GetDeclarations()
                                                           .Select(declaration => new { Key = kvp.Key, Declaration = declaration })
                               .Where(identifier => ((dynamic)identifier.Declaration).asTypeClause() == null));

            return targets.Select(target => new VariableTypeNotDeclaredInspectionResult(Name, target.Declaration, Severity, target.Key.ProjectName, target.Key.ModuleName));
        }
    }
}