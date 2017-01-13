using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNameInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public IdentifierNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parserState, IMessageBox messageBox, IPersistanceService<CodeInspectionSettings> settings)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new RenameDeclarationQuickFix(target.Context, target.QualifiedSelection, target, parserState, messageBox),
                new IgnoreOnceQuickFix(Context, target.QualifiedSelection, Inspection.AnnotationName), 
                new AddIdentifierToWhiteListQuickFix(Context, target.QualifiedSelection, target, settings)
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.IdentifierNameInspectionResultFormat, RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType, UI.Settings.Settings.Culture), Target.IdentifierName).Captialize(); }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
