using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ToDoListSettings : IToDoListSettings
    {
        [XmlArrayItem("ToDoMarker", IsNullable = false)]
        public ToDoMarker[] ToDoMarkers { get; set; }

        public ToDoListSettings()
        {
            var note = new ToDoMarker(RubberduckUI.TodoMarkerNote);
            var todo = new ToDoMarker(RubberduckUI.TodoMarkerTodo);
            var bug = new ToDoMarker(RubberduckUI.TodoMarkerBug);

            ToDoMarkers = new[] { note, todo, bug };
        }

        public ToDoListSettings(ToDoMarker[] markers)
        {
            ToDoMarkers = markers;
        }
    }
}
