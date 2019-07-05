using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.SmartIndenter
{
    public class IndenterConfigProvider : ConfigurationServiceBase<IndenterSettings>
    {
        public IndenterConfigProvider(IPersistenceService<IndenterSettings> persister)
            : base (persister, new FixedValueDefault<IndenterSettings>(new IndenterSettings(false)))
        { }
    }
}
