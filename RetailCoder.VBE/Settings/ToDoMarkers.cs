using System;
using System.Configuration;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IToDoMarker
    {
        string Text { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class ToDoMarker : IToDoMarker
    {
        //either the code can be properly case, or the XML can be, but the xml attributes must here *exactly* match the xml
        [XmlAttribute]
        public string Text { get; set; }
        
        [Obsolete]
        [XmlIgnore]
        public TodoPriority Priority { get; set; }
        
        /// <summary>   Default constructor is required for serialization. DO NOT USE. </summary>
        public ToDoMarker()
        {
            // default constructor required for serialization
        }

        public ToDoMarker(string text)
        {
            Text = text;
        }

        [Obsolete]
        public ToDoMarker(string text, TodoPriority priority) : this(text)
        {
        }

        /// <summary>   Convert this object into a string representation. Overriden for easy databinding.</summary>
        /// <returns>   The Text property. </returns>
        public override string ToString()
        {
            return Text;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ToDoMarker other))
            {
                return false;
            }

            return Text == other.Text;
        }

        public override int GetHashCode()
        {
            return Text.GetHashCode();
        }
    }

    public enum TodoPriority
    {
        Low, Medium, High
    }
}
