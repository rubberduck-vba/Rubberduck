using System.Collections.Generic;
using System.Globalization;
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
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly RubberduckParserState _parserState;
        private readonly IMessageBox _messageBox;
        private readonly IPersistanceService<CodeInspectionSettings> _settings;

        public IdentifierNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parserState, IMessageBox messageBox, 
                                              IPersistanceService<CodeInspectionSettings> settings)
            : base(inspection, target)
        {
            _parserState = parserState;
            _messageBox = messageBox;
            _settings = settings;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new RenameDeclarationQuickFix(Target.Context, Target.QualifiedSelection, Target, _parserState, _messageBox),
                    new IgnoreOnceQuickFix(Context, Target.QualifiedSelection, Inspection.AnnotationName), 
                    new AddIdentifierToWhiteListQuickFix(Context, Target.QualifiedSelection, Target, _settings)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.IdentifierNameInspectionResultFormat, RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType, CultureInfo.CurrentUICulture), Target.IdentifierName).Captialize(); }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
