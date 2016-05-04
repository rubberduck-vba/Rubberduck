using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
        }

        public override string Meta { get { return InspectionsUI.ImplicitActiveWorkbookReferenceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitActiveWorkbookReferenceInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly string[] Targets = 
        {
            "Worksheets", "Sheets", "Names", 
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (!_hostApp.IsValueCreated || _hostApp.Value == null || _hostApp.Value.ApplicationName != "Excel")
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
                new ImplicitActiveWorkbookReferenceInspectionResult(this, issue));
        }
    }

    public class ImplicitActiveWorkbookReferenceInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitActiveWorkbookReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
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