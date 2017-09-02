using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AddIdentifierToWhiteListQuickFix : QuickFixBase, IQuickFix
    {
        private readonly IPersistanceService<CodeInspectionSettings> _settings;

        public AddIdentifierToWhiteListQuickFix(IPersistanceService<CodeInspectionSettings> settings, InspectionLocator inspectionLocator)
        {
            _settings = settings;
            RegisterInspections(inspectionLocator.GetInspection<HungarianNotationInspection>(),
                inspectionLocator.GetInspection<UseMeaningfulNameInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var inspectionSettings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelist = inspectionSettings.WhitelistedIdentifiers;
            inspectionSettings.WhitelistedIdentifiers =
                whitelist.Concat(new[] { new WhitelistedIdentifierSetting(result.Target.IdentifierName) }).ToArray();
            _settings.Save(inspectionSettings);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.WhiteListIdentifierQuickFix;
        }

        public bool CanFixInProcedure { get; } = false;
        public bool CanFixInModule { get; } = false;
        public bool CanFixInProject { get; } = false;
    }
}
