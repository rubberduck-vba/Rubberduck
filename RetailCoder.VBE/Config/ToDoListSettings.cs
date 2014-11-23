using System.Xml.Serialization;
using System.Runtime.InteropServices;

namespace Rubberduck.Config
{
    interface IToDoListSettings
    {
        ToDoMarker[] ToDoMarkers { get; set; }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoListSettings : IToDoListSettings
    {
        [XmlArrayItemAttribute("ToDoMarker", IsNullable = false)]
        public ToDoMarker[] ToDoMarkers { get; set; }
    }
}
