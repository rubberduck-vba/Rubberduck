using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    internal interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
        ToDoExplorerColumnHeadingsOrder ColumnHeadingsOrder { get; set; }
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

        public ToDoExplorerColumnHeadingsOrder ColumnHeadingsOrder { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoListSettings()
        {
        }

        public ToDoListSettings(IEnumerable<ToDoMarker> defaultMarkers, ToDoExplorerColumnHeadingsOrder columnHeadingsOrder)
        {
            _markers = defaultMarkers;
            ColumnHeadingsOrder = columnHeadingsOrder;
        }

        public bool Equals(ToDoListSettings other)
        {
            return other != null 
                && ToDoMarkers.SequenceEqual(other.ToDoMarkers)
                && ColumnHeadingsOrder.Equals(other.ColumnHeadingsOrder);
        }
    }

    public class ToDoExplorerColumnHeadingsOrder : IEquatable<ToDoExplorerColumnHeadingsOrder>
    {
        public int DescriptionColumnIndex { get; set; }
        public int ProjectColumnIndex { get; set; }
        public int ModuleColumnIndex { get; set; }
        public int LineNumberColumnIndex { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoExplorerColumnHeadingsOrder()
        {
        }

        public ToDoExplorerColumnHeadingsOrder(int descriptionColumnIndex = 0, int projectColumnIndex = 1, int moduleColumnIndex = 2, int lineNumberColumnIndex = 3)
        {
            DescriptionColumnIndex = descriptionColumnIndex;
            ProjectColumnIndex = projectColumnIndex;
            ModuleColumnIndex = moduleColumnIndex;
            LineNumberColumnIndex = lineNumberColumnIndex;
        }

        public bool Equals(ToDoExplorerColumnHeadingsOrder other)
        {
            return DescriptionColumnIndex == other.DescriptionColumnIndex
                && ProjectColumnIndex == other.ProjectColumnIndex
                && ModuleColumnIndex == other.ModuleColumnIndex
                && LineNumberColumnIndex == other.LineNumberColumnIndex;
        }
    }
}
