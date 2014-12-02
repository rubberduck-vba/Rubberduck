using System.Xml.Serialization;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    public interface IToDoMarker
    {
        int Priority { get; set; }
        string Text { get; set; }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoMarker : IToDoMarker
    {
        //either the code can be properly case, or the XML can be, but the xml attributes must here *exactly* match the xml
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
