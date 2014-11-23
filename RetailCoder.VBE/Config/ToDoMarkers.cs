using System.Xml.Serialization;
using System.Runtime.InteropServices;

namespace Rubberduck.Config
{
    interface IToDoMarker
    {
        byte priority { get; set; }
        string text { get; set; }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoMarker : IToDoMarker
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
