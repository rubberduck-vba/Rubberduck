using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Parsing.Settings
{
    public class IgnoredProjectsConfigProvider : ConfigurationServiceBase<IgnoredProjectsSettings>
    {
        public IgnoredProjectsConfigProvider(IPersistenceService<IgnoredProjectsSettings> persister)
            : base(persister, new DefaultSettings<IgnoredProjectsSettings, Properties.ParsingSettings>())
        { }

        public override IgnoredProjectsSettings ReadDefaults()
        {
            return new IgnoredProjectsSettings();
        }
    }
}
