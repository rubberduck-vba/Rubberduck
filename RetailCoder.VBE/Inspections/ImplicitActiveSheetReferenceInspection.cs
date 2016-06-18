using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitActiveSheetReferenceInspection : InspectionBase
    {
        private readonly IHostApplication _hostApp;

        public ImplicitActiveSheetReferenceInspection(VBE vbe, RubberduckParserState state)
            : base(state)
        {
            _hostApp = vbe.HostApplication();
        }

        public override string Meta { get { return InspectionsUI.ImplicitActiveSheetReferenceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitActiveSheetReferenceInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly string[] Targets = 
        {
            "Cells", "Range", "Columns", "Rows"
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (_hostApp == null || _hostApp.ApplicationName != "Excel")
            {
                return new InspectionResultBase[] {};
                // if host isn't Excel, the ExcelObjectModel declarations shouldn't be loaded anyway.
            }

            var matches = BuiltInDeclarations.Where(item =>
                        Targets.Contains(item.IdentifierName) &&
                        item.ParentScope == "EXCEL.EXE;Excel._Global" &&
                        item.AsTypeName == "Range").ToList();

            var issues = matches.Where(item => item.References.Any())
                .SelectMany(declaration => declaration.References);

            return issues.Select(issue => 
                new ImplicitActiveSheetReferenceInspectionResult(this, issue));
        }
    }

    public class ImplicitActiveSheetReferenceInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitActiveSheetReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new IgnoreOnceQuickFix(reference.Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(Inspection.Description, _reference.Declaration.IdentifierName); }
        }
    }
}
