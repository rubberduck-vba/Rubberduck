using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    internal interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
        ToDoExplorerColumns ColumnHeaderInformation { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ToDoListSettings : IToDoListSettings, IEquatable<ToDoListSettings>
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

        public ToDoExplorerColumns ColumnHeaderInformation { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoListSettings()
        {
        }

        public ToDoListSettings(IEnumerable<ToDoMarker> defaultMarkers, ToDoExplorerColumns columnHeadings)
        {
            _markers = defaultMarkers;
            ColumnHeaderInformation = columnHeadings;
        }

        public bool Equals(ToDoListSettings other)
        {
            return other != null 
                && ToDoMarkers.SequenceEqual(other.ToDoMarkers)
                && ColumnHeaderInformation.Equals(other.ColumnHeaderInformation);
        }
    }

    public class ToDoExplorerColumns : IEquatable<ToDoExplorerColumns>
    {
        public ToDoExplorerColumn DescriptionColumn { get; set; }
        public ToDoExplorerColumn ProjectColumn { get; set; }
        public ToDoExplorerColumn ModuleColumn { get; set; }
        public ToDoExplorerColumn LineNumberColumn { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoExplorerColumns()
        {
        }

        public ToDoExplorerColumns(ToDoExplorerColumn descriptionColumn, ToDoExplorerColumn projectColumn, ToDoExplorerColumn moduleColumn, ToDoExplorerColumn lineNumberColumn)
        {
            DescriptionColumn = descriptionColumn;
            ProjectColumn = projectColumn;
            ModuleColumn = moduleColumn;
            LineNumberColumn = lineNumberColumn;
        }

        public bool Equals(ToDoExplorerColumns other)
        {
            return DescriptionColumn == other.DescriptionColumn 
                && ProjectColumn == other.ProjectColumn
                && ModuleColumn == other.ModuleColumn
                && LineNumberColumn == other.LineNumberColumn;
        }
    }

    public class ToDoExplorerColumn : IEquatable<ToDoExplorerColumn>
    {
        public int DisplayIndex { get; set; }
        [XmlElement(Type = typeof(DataGridLength))]
        public DataGridLength Width { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoExplorerColumn()
        {
        }

        public ToDoExplorerColumn(int displayIndex, DataGridLength width)
        {
            DisplayIndex = displayIndex;
            Width = width;
        }

        public bool Equals(ToDoExplorerColumn other)
        {
            return DisplayIndex == other.DisplayIndex
                && Width == other.Width;
        }
    }
}
