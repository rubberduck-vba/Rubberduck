using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class UseMeaningfulNameInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public UseMeaningfulNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parserState, IMessageBox messageBox)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new RenameDeclarationQuickFix(target.Context, target.QualifiedSelection, target, parserState, messageBox),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.UseMeaningfulNameInspectionResultFormat, RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType, UI.Settings.Settings.Culture), Target.IdentifierName).Captialize(); }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
