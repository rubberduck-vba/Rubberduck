using System.Collections.Generic;
using Castle.DynamicProxy.Generators;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class MoveFieldCloserToUsageInspectionResult : InspectionResultBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private IEnumerable<QuickFixBase> _quickFixes;

        public MoveFieldCloserToUsageInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, target)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new MoveFieldCloserToUsageQuickFix(Target.Context, Target.QualifiedSelection, Target, _state, _messageBox),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }
    }
}
