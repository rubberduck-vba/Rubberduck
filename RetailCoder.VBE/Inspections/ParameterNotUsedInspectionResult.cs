using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation, 
            VBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = isInterfaceImplementation ? new CodeInspectionQuickFix[] {} : new CodeInspectionQuickFix[]
            {
                new RemoveUnusedParameterQuickFix(Context, QualifiedSelection, vbe, state, messageBox),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterNotUsedInspectionResultFormat, Target.IdentifierName); }
        }
    }

    public class RemoveUnusedParameterQuickFix : CodeInspectionQuickFix
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RemoveUnusedParameterQuickFix(ParserRuleContext context, QualifiedSelection selection, 
            VBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(context, selection, InspectionsUI.RemoveUnusedParameterQuickFix)
        {
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix()
        {
            using (var dialog = new RemoveParametersDialog())
            {
                var refactoring = new RemoveParametersRefactoring(_vbe,
                    new RemoveParametersPresenterFactory(_vbe, dialog, _state, _messageBox));

                refactoring.QuickFix(_state, Selection);
            }
        }
    }
}
