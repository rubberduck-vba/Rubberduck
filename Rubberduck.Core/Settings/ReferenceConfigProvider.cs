using System;
using System.Reflection;
using Rubberduck.Resources.Registration;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;

namespace Rubberduck.Settings
{
    public class ReferenceConfigProvider : IConfigProvider<ReferenceSettings>
    {
        private readonly IPersistanceService<ReferenceSettings> _persister;

        public ReferenceConfigProvider(IPersistanceService<ReferenceSettings> persister)
        {
            _persister = persister;
        }

        public ReferenceSettings Create()
        {
            var defaults = CreateDefaults();
            return _persister.Load(defaults) ?? defaults;
        }

        public ReferenceSettings CreateDefaults()
        {
            var defaults = new ReferenceSettings
            {
                RecentReferencesTracked = 20
            };

            var version = Assembly.GetEntryAssembly().GetName().Version;
            defaults.PinReference(new ReferenceInfo(new Guid(RubberduckGuid.RubberduckTypeLibGuid), string.Empty, string.Empty, version.Major, version.Minor));
            defaults.PinReference(new ReferenceInfo(new Guid(RubberduckGuid.RubberduckApiTypeLibGuid), string.Empty, string.Empty, version.Major, version.Minor));

            return defaults;
        }

        public void Save(ReferenceSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
