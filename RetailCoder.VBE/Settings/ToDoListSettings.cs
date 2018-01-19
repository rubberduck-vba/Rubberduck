using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    internal interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
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

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoListSettings()
        {
        }

        public ToDoListSettings(IEnumerable<ToDoMarker> defaultMarkers)
        {
            _markers = defaultMarkers;
        }

        public bool Equals(ToDoListSettings other)
        {
            return other != null && ToDoMarkers.SequenceEqual(other.ToDoMarkers);
        }
    }
}
