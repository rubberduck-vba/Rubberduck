using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IDictionary<QualifiedModuleName, IParseTree> nodes)
        {
            foreach (var module in nodes)
            {
                var options = module.Value.GetModuleOptions().Select(
                    option => new {Key = module.Key, Option = option}).ToList();

                if (!options.Any() || options.All(option => option.Option.children.Last().GetText() != ReservedKeywords.Explicit))
                {
                    yield return new OptionExplicitInspectionResult(Name, null, Severity, module.Key.ProjectName, module.Key.ModuleName);
                }
            }
        }
    }
}