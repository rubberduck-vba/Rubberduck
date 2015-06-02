using System.Xml.Serialization;

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
            //empty constructor needed for serialization
        }

        public ToDoListSettings(ToDoMarker[] markers)
        {
            this.ToDoMarkers = markers;
        }
    }
}
