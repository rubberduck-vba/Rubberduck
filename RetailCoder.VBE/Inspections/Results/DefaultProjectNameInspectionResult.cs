using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class DefaultProjectNameInspectionResult : InspectionResultBase
    {
        private IEnumerable<IQuickFix> _quickFixes;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, target)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new RenameProjectQuickFix(Target.Context, Target.QualifiedSelection, Target, _state, _messageBox)
                });
            }
        }

        public override string Description
        {
            get { return Inspection.Description.Captialize(); }
        }
    }
}
