using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitActiveSheetReferenceInspection : InspectionBase
    {
        private readonly IHostApplication _hostApp;

        public ImplicitActiveSheetReferenceInspection(IVBE vbe, RubberduckParserState state)
            : base(state)
        {
            _hostApp = vbe.HostApplication();
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
            if (_hostApp == null || _hostApp.ApplicationName != "Excel")
            {
                return Enumerable.Empty<InspectionResultBase>();
                // if host isn't Excel, the ExcelObjectModel declarations shouldn't be loaded anyway.
            }

            var matches = BuiltInDeclarations.Where(item =>
                        item.ProjectName == "Excel" &&
                        Targets.Contains(item.IdentifierName) &&
                        item.ParentDeclaration.ComponentName == "Global" &&
                        item.AsTypeName == "Range").ToList();

            var issues = matches.Where(item => item.References.Any())
                .SelectMany(declaration => declaration.References.Distinct());

            return issues.Select(issue => 
                new ImplicitActiveSheetReferenceInspectionResult(this, issue));
        }
    }
}
