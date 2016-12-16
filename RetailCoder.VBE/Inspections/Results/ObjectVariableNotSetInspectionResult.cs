using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public sealed class ObjectVariableNotSetInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ObjectVariableNotSetInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new QuickFixBase[]
            {
                new UseSetKeywordForObjectAssignmentQuickFix(_reference),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName),
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ObjectVariableNotSetInspectionResultFormat, _reference.Declaration.IdentifierName).Captialize(); }
        }
    }
}