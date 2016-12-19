using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Results
{
    public class DefaultProjectNameInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes; 

        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state)
            : base(inspection, target)
        {
            _quickFixes = new[]
            {
                new RenameProjectQuickFix(target.Context, target.QualifiedSelection, target, state),
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return Inspection.Description.Captialize(); }
        }
    }
}
