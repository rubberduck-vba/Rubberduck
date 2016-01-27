using System;
using System.Xml.Serialization;
using Rubberduck.ToDoItems;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    public interface IToDoMarker
    {
        TodoPriority Priority { get; set; }
        string Text { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class ToDoMarker : IToDoMarker
    {
        //either the code can be properly case, or the XML can be, but the xml attributes must here *exactly* match the xml
        [XmlAttribute]
        public string Text { get; set; }

        [XmlAttribute]
        public TodoPriority Priority { get; set; }

        [XmlIgnore]
        public string PriorityLabel
        {
            get { return RubberduckUI.ResourceManager.GetString("ToDoPriority_" + Priority, RubberduckUI.Culture); }
            set
            {
                foreach (var priority in Enum.GetValues(typeof(TodoPriority)))
                {
                    if (value == RubberduckUI.ResourceManager.GetString("ToDoPriority_" + priority, RubberduckUI.Culture))
                    {
                        Priority = (TodoPriority)priority;
                        return;
                    }
                }
            }
        }

        /// <summary>   Default constructor is required for serialization. DO NOT USE. </summary>
        public ToDoMarker()
        {
            // default constructor required for serialization
        }

        public ToDoMarker(string text, TodoPriority priority)
        {
            Text = text;
            Priority = priority;
        }

        /// <summary>   Convert this object into a string representation. Over-riden for easy databinding.</summary>
        /// <returns>   The Text property. </returns>
        public override string ToString()
        {
            return Text;
        }

        public override bool Equals(object obj)
        {
            var other = obj as ToDoMarker;

            if (other == null) { return false; }

            // no need to check PriorityLabel as it soley relies on Priority - if one is wrong, the other has to be too
            return Text == other.Text &&
                   Priority == other.Priority;
        }

        public override int GetHashCode()
        {
            return Text.GetHashCode();
        }
    }
}
