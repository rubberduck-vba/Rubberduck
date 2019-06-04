using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Controls;
using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    internal interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
        ObservableCollection<GridViewColumnInfo> ColumnHeadersInformation { get; set; }
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

        private ObservableCollection<GridViewColumnInfo> _columnHeadersInfo;
        public ObservableCollection<GridViewColumnInfo> ColumnHeadersInformation
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

        public ToDoListSettings(IEnumerable<ToDoMarker> defaultMarkers, ObservableCollection<GridViewColumnInfo> columnHeaders)
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

    public class GridViewColumnInfo : ViewModelBase, IEquatable<GridViewColumnInfo>
    {
        private int _displayIndex;
        public int DisplayIndex
        {
            get => _displayIndex;
            set
            {
                if (value != _displayIndex)
                {
                    _displayIndex = value;
                    OnPropertyChanged();
                }
            }
        }

        [XmlElement(Type = typeof(DataGridLength))]
        private DataGridLength _width;

        public DataGridLength Width
        {
            get => _width;
            set
            {
                if (value != _width)
                {
                    _width = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public GridViewColumnInfo()
        {
        }

        public GridViewColumnInfo(int displayIndex, DataGridLength width)
        {
            DisplayIndex = displayIndex;
            Width = width;
        }

        public bool Equals(GridViewColumnInfo other)
        {
            return DisplayIndex == other.DisplayIndex
                && Width == other.Width;
        }
    }
}
