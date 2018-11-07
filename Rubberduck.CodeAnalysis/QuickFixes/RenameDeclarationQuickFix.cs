using System.Globalization;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RenameDeclarationQuickFix : QuickFixBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RenameDeclarationQuickFix(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(typeof(HungarianNotationInspection), typeof(UseMeaningfulNameInspection), typeof(DefaultProjectNameInspection), typeof(UnderscoreInPublicClassModuleMemberInspection))
        {
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix(IInspectionResult result)
        {
            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(_vbe, view, _state);
                var refactoring = new RenameRefactoring(_vbe, factory, _messageBox, _state);
                refactoring.Refactor(result.Target);
            }
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(RubberduckUI.Rename_DeclarationType,
                RubberduckUI.ResourceManager.GetString("DeclarationType_" + result.Target.DeclarationType,
                    CultureInfo.CurrentUICulture));
        }

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}