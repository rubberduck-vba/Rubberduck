using System.Xml.Serialization;

namespace Rubberduck.Config
{
    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoListSettings
    {

        [XmlArrayItemAttribute("ToDoMarker", IsNullable = false)]
        public ToDoMarker[] ToDoMarkers { get; set; }
    }

    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoMarker
    {
        [XmlAttributeAttribute()]
        public string text { get; set; }

        [XmlAttributeAttribute()]
        public byte priority { get; set; }

        public ToDoMarker() { }

        public ToDoMarker(string text, byte priority)
        {
            this.text = text;
            this.priority = priority;
        }
    }
}
