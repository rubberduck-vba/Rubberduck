using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class IgnoredProjectsConfigProvider : ConfigurationServiceBase<IgnoredProjectsSettings>
    {
        public IgnoredProjectsConfigProvider(IPersistenceService<IgnoredProjectsSettings> persister)
            : base(persister, new DefaultSettings<IgnoredProjectsSettings, Properties.Settings>())
        { }

        public override IgnoredProjectsSettings ReadDefaults()
        {
            return new IgnoredProjectsSettings();
        }
    }
}
