using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitActiveSheetReferenceInspection : InspectionBase
    {
        private readonly Func<IHostApplication> _hostApp;

        public ImplicitActiveSheetReferenceInspection(VBE vbe, RubberduckParserState state)
            : base(state)
        {
            _hostApp = vbe.HostApplication;
        }

        public override string Meta { get { return InspectionsUI.ImplicitActiveSheetReferenceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitActiveSheetReferenceInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly string[] Targets = 
        {
            "Cells", "Range", "Columns", "Rows"
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (_hostApp().ApplicationName != "Excel")
            {
                return new InspectionResultBase[] {};
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

    public class ImplicitActiveSheetReferenceInspectionResult : InspectionResultBase
    {
        private readonly string _result;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitActiveSheetReferenceInspectionResult(IInspection inspection, string result, ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context)
        {
            _result = result;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new IgnoreOnceQuickFix(context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return _result; }
        }
    }
}
