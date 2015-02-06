using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VbModuleParseResult> parseResult)
        {
            foreach (var module in parseResult)
            {
                var options = module.ParseTree.GetModuleOptions().ToList();

                if (!options.Any() || options.All(option => option.children.Last().GetText() != Tokens.Explicit))
                {
                    yield return new OptionExplicitInspectionResult(Name, Severity, module.QualifiedName);
                }
            }
        }
    }
}