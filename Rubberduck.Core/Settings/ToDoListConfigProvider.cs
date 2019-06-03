using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class ToDoListConfigProvider : ConfigurationServiceBase<ToDoListSettings>
    {
        private readonly IEnumerable<ToDoMarker> _defaultMarkers;
        private readonly ToDoExplorerColumns _columnHeadingsOrder;

        public ToDoListConfigProvider(IPersistenceService<ToDoListSettings> persister)
            : base(persister, new DefaultSettings<ToDoListSettings, Properties.Settings>())
        {
            _defaultMarkers = new DefaultSettings<ToDoMarker, Properties.Settings>().Defaults;
            _columnHeadingsOrder = new DefaultSettings<ToDoExplorerColumns, Properties.Settings>().Default;
        }
        
        public override ToDoListSettings ReadDefaults()
        {
            return new ToDoListSettings(_defaultMarkers, _columnHeadingsOrder);
        }
    }
}
