using System.Xml.Serialization;

namespace Rubberduck.Config
{
    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoListSettings
    {

        private ToDoMarker[] toDoMarkersField;

        [XmlArrayItemAttribute("ToDoMarker", IsNullable = false)]
        public ToDoMarker[] ToDoMarkers
        {
            get
            {
                return this.toDoMarkersField;
            }
            set
            {
                this.toDoMarkersField = value;
            }
        }
    }

    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class ToDoMarker
    {

        private string textField;
        private byte priorityField;

        [XmlAttributeAttribute()]
        public string text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }

        [XmlAttributeAttribute()]
        public byte priority
        {
            get
            {
                return this.priorityField;
            }
            set
            {
                this.priorityField = value;
            }
        }
    }
}
