using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Inspections.Results
{
    public class EncapsulatePublicFieldInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public EncapsulatePublicFieldInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state, IIndenter indenter)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new EncapsulateFieldQuickFix(target.Context, target.QualifiedSelection, target, state, indenter),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.EncapsulatePublicFieldInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
