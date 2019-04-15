using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class ToDoListConfigProvider : ConfigurationServiceBase<ToDoListSettings>
    {
        private readonly IEnumerable<ToDoMarker> defaultMarkers;

        public ToDoListConfigProvider(IPersistenceService<ToDoListSettings> persister)
            : base(persister, new DefaultSettings<ToDoListSettings, Properties.Settings>())
        {
            defaultMarkers = new DefaultSettings<ToDoMarker, Properties.Settings>().Defaults;
        }
        
        public override ToDoListSettings ReadDefaults()
        {
            return new ToDoListSettings(defaultMarkers);
        }
    }
}
