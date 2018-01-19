using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class ToDoListConfigProvider : IConfigProvider<ToDoListSettings>
    {
        private readonly IPersistanceService<ToDoListSettings> _persister;
        private readonly IEnumerable<ToDoMarker> _defaultMarkers;

        public ToDoListConfigProvider(IPersistanceService<ToDoListSettings> persister)
        {
            _persister = persister;
            _defaultMarkers = new DefaultSettings<ToDoMarker>().Defaults;
        }

        public ToDoListSettings Create()
        {
            var prototype = new ToDoListSettings(_defaultMarkers);
            return _persister.Load(prototype) ?? prototype;
        }

        public ToDoListSettings CreateDefaults()
        {
            return new ToDoListSettings(_defaultMarkers);
        }

        public void Save(ToDoListSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
