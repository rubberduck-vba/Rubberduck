using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class ToDoListConfigProvider : ConfigurationServiceBase<ToDoListSettings>
    {
        private readonly IEnumerable<ToDoMarker> _defaultMarkers;
        private readonly ObservableCollection<ToDoGridViewColumnInfo> _toDoExplorerColumns;

        public ToDoListConfigProvider(IPersistenceService<ToDoListSettings> persister)
            : base(persister, new DefaultSettings<ToDoListSettings, Properties.Settings>())
        {
            _defaultMarkers = new DefaultSettings<ToDoMarker, Properties.Settings>().Defaults;

            var gvciDefaults = new DefaultSettings<ToDoGridViewColumnInfo, Properties.Settings>().Defaults;
            _toDoExplorerColumns = new ObservableCollection<ToDoGridViewColumnInfo>(gvciDefaults);
        }
        
        public override ToDoListSettings ReadDefaults()
        {
            return new ToDoListSettings(_defaultMarkers, _toDoExplorerColumns);
        }

        public override ToDoListSettings Read()
        {
            var toDoListSettings = base.Read();

            if (toDoListSettings.ColumnHeadersInformation == null
                || toDoListSettings.ColumnHeadersInformation.Count == 0)
            {
                toDoListSettings.ColumnHeadersInformation = _toDoExplorerColumns;
            }

            return toDoListSettings;
        }
    }
}
