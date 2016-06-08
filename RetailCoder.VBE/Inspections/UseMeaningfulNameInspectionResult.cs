using System.Collections.Generic;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UseMeaningfulNameInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UseMeaningfulNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parserState, IMessageBox messageBox)
            : base(inspection, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RenameDeclarationQuickFix(target.Context, target.QualifiedSelection, target, parserState, messageBox),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.UseMeaningfulNameInspectionResultFormat, RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType), Target.IdentifierName); }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }

    /// <summary>
    /// A code inspection quickfix that addresses a VBProject bearing the default name.
    /// </summary>
    public class RenameDeclarationQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RenameDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(context, selection, string.Format(RubberduckUI.Rename_DeclarationType, RubberduckUI.ResourceManager.GetString("DeclarationType_" + target.DeclarationType, RubberduckUI.Culture)))
        {
            _target = target;
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(vbe, view, _state, _messageBox);
                var refactoring = new RenameRefactoring(vbe, factory, _messageBox, _state);
                refactoring.Refactor(_target);
                IsCancelled = view.DialogResult == DialogResult.Cancel;
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}
