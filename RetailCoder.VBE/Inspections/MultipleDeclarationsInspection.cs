using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class MultipleDeclarationsInspection : IInspection
    {
        public MultipleDeclarationsInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.MultipleDeclarations; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IDictionary<QualifiedModuleName, IParseTree> nodes)
        {
            var declarations = nodes.SelectMany(
                node => node.Value.GetDeclarations().Select(declaration => new {Key = node.Key, Declaration = declaration}))
                .Select(node => new {Key = node.Key, node.Declaration});

            return 
                declarations.Select(declaration => new MultipleDeclarationsInspectionResult(Name, declaration.Declaration, Severity,declaration.Key.ProjectName, declaration.Key.ModuleName)); 
        }
    }
}