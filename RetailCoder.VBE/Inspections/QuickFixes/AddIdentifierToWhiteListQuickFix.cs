using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AddIdentifierToWhiteListQuickFix : QuickFixBase
    {
        private readonly IPersistanceService<CodeInspectionSettings> _settings;
        private readonly Declaration _target;

        public AddIdentifierToWhiteListQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, IPersistanceService<CodeInspectionSettings> settings)
            : base(context, selection, InspectionsUI.WhiteListIdentifierQuickFix)
        {
            _settings = settings;
            _target = target;
        }

        public override void Fix()
        {
            var inspectionSettings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelist = inspectionSettings.WhitelistedIdentifiers;
            inspectionSettings.WhitelistedIdentifiers =
                whitelist.Concat(new[] { new WhitelistedIdentifierSetting(_target.IdentifierName) }).ToArray();
            _settings.Save(inspectionSettings);
        }
    }
}
