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
            //_toDoExplorerColumns = new DefaultSettings<ObservableCollection<GridViewColumnInfo>, Properties.Settings>().Default;
            //TODO: Clean up :barf:. Deserialization as the ^ `DefaultSettings<T,U>()Default` is null.
            _toDoExplorerColumns = new ObservableCollection<ToDoGridViewColumnInfo>
            {
                new ToDoGridViewColumnInfo(0, new DataGridLength(1, DataGridLengthUnitType.Auto)),
                new ToDoGridViewColumnInfo(1, new DataGridLength(75)),
                new ToDoGridViewColumnInfo(2, new DataGridLength(75)),
                new ToDoGridViewColumnInfo(3, new DataGridLength(75))
            };
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
