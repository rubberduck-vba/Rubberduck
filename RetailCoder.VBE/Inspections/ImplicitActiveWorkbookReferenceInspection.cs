using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.UI;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Inspections
{
    public class ImplicitActiveWorkbookReferenceInspection : IInspection
    {
        private readonly Lazy<IHostApplication> _hostApp;

        public ImplicitActiveWorkbookReferenceInspection(VBE vbe)
        {
            _hostApp = new Lazy<IHostApplication>(vbe.HostApplication);
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ImplicitActiveWorkbookReferenceInspection"; } }
        public string Description { get { return RubberduckUI.ImplicitActiveWorkbookReference_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly string[] Targets = 
        {
            "Worksheets", "Sheets", "Names", 
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
                                                                      && item.References.Any())
                .SelectMany(declaration => declaration.References);

            return issues.Select(issue => 
                new ImplicitActiveSheetReferenceInspectionResult(this, string.Format(Description, issue.Declaration.IdentifierName), issue.Context, issue.QualifiedModuleName));
        }
    }
}