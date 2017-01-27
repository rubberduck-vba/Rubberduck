using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    interface IToDoListSettings
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
            get { return _markers.ToArray(); }
            set
            {
                //Only take the first marker if there are duplicates.
                _markers = value.GroupBy(m => m.Text).Select(marker => marker.First()).ToArray();
            }
        }

        public ToDoListSettings()
        {
            var note = new ToDoMarker(RubberduckUI.TodoMarkerNote);
            var todo = new ToDoMarker(RubberduckUI.TodoMarkerTodo);
            var bug = new ToDoMarker(RubberduckUI.TodoMarkerBug);

            ToDoMarkers = new[] { note, todo, bug };
        }

        public ToDoListSettings(IEnumerable<ToDoMarker> markers)
        {
            _markers = markers;
        }

        public bool Equals(ToDoListSettings other)
        {
            return other != null && ToDoMarkers.SequenceEqual(other.ToDoMarkers);
        }
    }
}
