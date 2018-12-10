using System;
using System.IO;
using System.Reflection;
using Rubberduck.Resources.Registration;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Settings
{
    public class ReferenceConfigProvider : IConfigProvider<ReferenceSettings>
    {
        private readonly IPersistanceService<ReferenceSettings> _persister;
        private readonly IEnvironmentProvider _environment;

        public ReferenceConfigProvider(IPersistanceService<ReferenceSettings> persister, IEnvironmentProvider environment)
        {
            _persister = persister;
            _environment = environment;
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

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            defaults.PinReference(new ReferenceInfo(new Guid(RubberduckGuid.RubberduckTypeLibGuid), string.Empty, string.Empty, version.Major, version.Minor));
            defaults.PinReference(new ReferenceInfo(new Guid(RubberduckGuid.RubberduckApiTypeLibGuid), string.Empty, string.Empty, version.Major, version.Minor));

            var documents = _environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (!string.IsNullOrEmpty(documents))
            {
                defaults.ProjectPaths.Add(documents);
            }

            var appdata = _environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (!string.IsNullOrEmpty(documents))
            {
                var addins = Path.Combine(appdata, "Microsoft", "AddIns");
                if (Directory.Exists(addins))
                {
                    defaults.ProjectPaths.Add(addins);
                }

                var templates = Path.Combine(appdata, "Microsoft", "Templates");
                if (Directory.Exists(templates))
                {
                    defaults.ProjectPaths.Add(templates);
                }
            }

            return defaults;
        }

        public void Save(ReferenceSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
