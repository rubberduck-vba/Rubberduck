using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitActiveWorkbookReferenceInspection : InspectionBase
    {
        private readonly Lazy<IHostApplication> _hostApp;

        public ImplicitActiveWorkbookReferenceInspection(VBE vbe, RubberduckParserState state)
            : base(state)
        {
            _hostApp = new Lazy<IHostApplication>(vbe.HostApplication);
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.ImplicitActiveWorkbookReference_; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly string[] Targets = 
        {
            "Worksheets", "Sheets", "Names", 
        };

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            if (!_hostApp.IsValueCreated || _hostApp.Value == null || _hostApp.Value.ApplicationName != "Excel")
            {
                return new CodeInspectionResultBase[] {};
                // if host isn't Excel, the ExcelObjectModel declarations shouldn't be loaded anyway.
            }

            var issues = Declarations.Where(item => item.IsBuiltIn 
                                            && item.ParentScope == "Excel.Global"
                                            && Targets.Contains(item.IdentifierName)
                                            && item.References.Any())
                .SelectMany(declaration => declaration.References);

            return issues.Select(issue => 
                new ImplicitActiveSheetReferenceInspectionResult(this, string.Format(Description, issue.Declaration.IdentifierName), issue.Context, issue.QualifiedModuleName));
        }
    }
}