using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
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
        public string Description { get { return RubberduckUI.ImplicitActiveSheetReference_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly string[] Targets = 
        {
            "Cells", "Range", "Columns", "Rows"
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            if (_hostApp.Value.ApplicationName != "Excel")
            {
                return new CodeInspectionResultBase[] {};
                // if host isn't Excel, the ExcelObjectModel declarations shouldn't be loaded anyway.
            }

            var issues = parseResult.AllDeclarations.Where(item => item.IsBuiltIn 
                && item.ParentScope == "Excel.Global"
                && Targets.Contains(item.IdentifierName)
                && item.References.Any())
                .SelectMany(declaration => declaration.References);

            return issues.Select(issue => 
                new ImplicitActiveSheetReferenceInspectionResult(this, string.Format(Description, issue.Declaration.IdentifierName), issue.Context, issue.QualifiedModuleName));
        }
    }

    public class ImplicitActiveSheetReferenceInspectionResult : CodeInspectionResultBase
    {
        //private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitActiveSheetReferenceInspectionResult(IInspection inspection, string result, ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, result, qualifiedName, context)
        {
            //_quickFixes = new CodeInspectionQuickFix[]{};
        }

        //public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }
}
