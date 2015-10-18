using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Inspections
{
    public class ImplicitActiveSheetReferenceInspection : IInspection
    {
        private readonly Lazy<IHostApplication> _hostApp;

        public ImplicitActiveSheetReferenceInspection(VBE vbe)
        {
            _hostApp = new Lazy<IHostApplication>(vbe.HostApplication);
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ImplicitActiveSheetReferenceInspection"; } }
        public string Description { get { return RubberduckUI.ImplicitActiveSheetReference; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly string[] Targets = 
        {
            "Selection", "Cells", "Range", "Names", "Columns", "Rows"
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            if (_hostApp.Value.ApplicationName != "Excel")
            {
                return new CodeInspectionResultBase[] {};
                // if host isn't Excel, the ExcelObjectModel declarations shouldn't be loaded anyway.
            }

            var issues = parseResult.Declarations.Items.Where(item => item.IsBuiltIn 
                && item.ParentScope.StartsWith("Excel.Global")
                && Targets.Contains(item.IdentifierName)
                && item.References.Any());

            return new CodeInspectionResultBase[] { }; // todo: return an inspection result for each item in 'issues'
        }
    }
}
