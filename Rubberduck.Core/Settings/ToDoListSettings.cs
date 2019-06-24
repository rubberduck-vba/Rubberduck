using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    internal interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
        ObservableCollection<ToDoGridViewColumnInfo> ColumnHeadersInformation { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ToDoListSettings : ViewModelBase, IToDoListSettings, IEquatable<ToDoListSettings>
    {
        private IEnumerable<ToDoMarker> _markers;

        [XmlArrayItem("ToDoMarker", IsNullable = false)]
        public ToDoMarker[] ToDoMarkers
        {
            get => _markers.ToArray();
            set
            {
                //Only take the first marker if there are duplicates.
                _markers = value.GroupBy(m => m.Text).Select(marker => marker.First()).ToArray();
            }
        }

        private ObservableCollection<ToDoGridViewColumnInfo> _columnHeadersInfo;
        public ObservableCollection<ToDoGridViewColumnInfo> ColumnHeadersInformation
        {
            get => _columnHeadersInfo;
            set
            {
                if (value != _columnHeadersInfo)
                {
                    _columnHeadersInfo = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoListSettings()
        {
        }

        public ToDoListSettings(IEnumerable<ToDoMarker> defaultMarkers, ObservableCollection<ToDoGridViewColumnInfo> columnHeaders)
        {
            _markers = defaultMarkers;
            ColumnHeadersInformation = columnHeaders;
        }

        public bool Equals(ToDoListSettings other)
        {
            return other != null 
                && ToDoMarkers.SequenceEqual(other.ToDoMarkers)
                && ColumnHeadersInformation.Equals(other.ColumnHeadersInformation);
        }
    }
}
