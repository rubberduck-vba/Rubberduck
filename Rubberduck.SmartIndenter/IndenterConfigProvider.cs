using Rubberduck.SettingsProvider;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenterConfigProvider
    {
        IndenterSettings Create();
        IndenterSettings CreateDefaults();

        void Save(IndenterSettings settings);
    }

    public class IndenterConfigProvider : IIndenterConfigProvider
    {
        private readonly IPersistanceService<IndenterSettings> _persister;

        public IndenterConfigProvider(IPersistanceService<IndenterSettings> persister)
        {
            _persister = persister;
        }

        public IndenterSettings Create()
        {
            var prototype = new IndenterSettings();
            return _persister.Load(prototype) ?? prototype;
        }

        public IndenterSettings CreateDefaults()
        {
            return new IndenterSettings();
        }

        public void Save(IndenterSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
