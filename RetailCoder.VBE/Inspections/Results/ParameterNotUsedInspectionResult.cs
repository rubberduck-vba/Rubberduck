using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.Results
{
    public class ParameterNotUsedInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly bool _isInterfaceImplementation;
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public ParameterNotUsedInspectionResult(IInspection inspection, Declaration target,
            bool isInterfaceImplementation, IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, target)
        {
            _isInterfaceImplementation = isInterfaceImplementation;
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _isInterfaceImplementation
                    ? new QuickFixBase[] { }
                    : (_quickFixes ?? (_quickFixes = new QuickFixBase[]
                    {
                        new RemoveUnusedParameterQuickFix(Context, QualifiedSelection, _vbe, _state, _messageBox),
                        new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                    }));
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterNotUsedInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
