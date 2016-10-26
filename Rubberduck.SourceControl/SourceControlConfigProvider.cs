using System;
using System.IO;
using Rubberduck.SettingsProvider;

namespace Rubberduck.SourceControl
{
    public class SourceControlConfigProvider : IConfigProvider<SourceControlSettings>
    {
        private readonly string _rootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");
        private readonly IFilePersistanceService<SourceControlSettings> _persister;

        public SourceControlConfigProvider(IFilePersistanceService<SourceControlSettings> persister)
        {
            _persister = persister;
            _persister.FilePath = Path.Combine(_rootPath, "SourceControl.rubberduck");
        }

        public SourceControlSettings Create()
        {
            var prototype = new SourceControlSettings();
            return _persister.Load(prototype) ?? prototype;
        }

        public SourceControlSettings CreateDefaults()
        {
            return new SourceControlSettings();
        }

        public void Save(SourceControlSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
