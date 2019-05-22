using System.Collections.Generic;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class ToDoListConfigProvider : ConfigurationServiceBase<ToDoListSettings>
    {
        private readonly IEnumerable<ToDoMarker> _defaultMarkers;
        private readonly ToDoExplorerColumnHeadingsOrder _columnHeadingsOrder;

        public ToDoListConfigProvider(IPersistenceService<ToDoListSettings> persister)
            : base(persister, new DefaultSettings<ToDoListSettings, Properties.Settings>())
        {
            _defaultMarkers = new DefaultSettings<ToDoMarker, Properties.Settings>().Defaults;

            // TODO: Figure out how to add defaults
            //_columnHeadingsOrder = new ToDoExplorerColumnHeadingsOrder(3, 2, 1, 0);
            _columnHeadingsOrder = new DefaultSettings<ToDoExplorerColumnHeadingsOrder, Properties.Settings>().Default;
        }
        
        public override ToDoListSettings ReadDefaults()
        {
            return new ToDoListSettings(_defaultMarkers, _columnHeadingsOrder);
        }
    }
}
