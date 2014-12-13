using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class OptionExplicitInspection : IInspection
    {
        public OptionExplicitInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.OptionExplicit; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(SyntaxTreeNode node)
        {
            foreach (var module in node.ChildNodes.OfType<ModuleNode>())
            {
                var options = module.ChildNodes.OfType<OptionNode>().ToList();
                if (!options.Any() || options.All(option => option.Option != ReservedKeywords.Explicit))
                {
                    yield return new OptionExplicitInspectionResult(Name, module.Instruction, Severity);
                }
            }
        }
    }
}