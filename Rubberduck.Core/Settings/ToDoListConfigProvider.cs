using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class ToDoListConfigProvider : ConfigurationServiceBase<ToDoListSettings>
    {
        private readonly IEnumerable<ToDoMarker> defaultMarkers;

        public ToDoListConfigProvider(IPersistanceService<ToDoListSettings> persister)
            : base(persister)
        {
            defaultMarkers = new DefaultSettings<ToDoMarker, Properties.Settings>().Defaults;
        }

        public override ToDoListSettings Load()
        {
            var prototype = new ToDoListSettings(defaultMarkers);
            return persister.Load(prototype) ?? prototype;
        }

        public override ToDoListSettings LoadDefaults()
        {
            return new ToDoListSettings(defaultMarkers);
        }
    }
}
