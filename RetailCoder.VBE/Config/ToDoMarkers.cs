using System.Xml.Serialization;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Config
{
    interface IToDoMarker
    {
        int Priority { get; set; }
        string Text { get; set; }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoMarker : IToDoMarker
    {
        [XmlAttribute]
        public string Text { get; set; }

        [XmlAttribute]
        public int Priority { get; set; }

        public ToDoMarker()
        {
            // default constructor required for serialization
        }

        public ToDoMarker(string text, int priority)
        {
            Text = text;
            Priority = priority;
        }
    }
}
