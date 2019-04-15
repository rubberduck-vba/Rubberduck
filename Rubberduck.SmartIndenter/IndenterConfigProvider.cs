using System;
using Rubberduck.SettingsProvider;

namespace Rubberduck.SmartIndenter
{
    public class IndenterConfigProvider : IConfigurationService<IndenterSettings>
    {
        private readonly IPersistanceService<IndenterSettings> _persister;

        public IndenterConfigProvider(IPersistanceService<IndenterSettings> persister)
        {
            _persister = persister;
        }

        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;

        public IndenterSettings Read()
        {
            var prototype = new IndenterSettings(false);
            return _persister.Load(prototype) ?? prototype;
        }

        public IndenterSettings ReadDefaults()
        {
            return new IndenterSettings(false);
        }

        public void Save(IndenterSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
