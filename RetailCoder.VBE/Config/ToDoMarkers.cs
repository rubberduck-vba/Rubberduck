﻿using System.Xml.Serialization;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    public enum TodoPriority
    {
        Low, 
        Normal,
        High
    }

    [ComVisible(false)]
    public interface IToDoMarker
    {
        TodoPriority Priority { get; set; }
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
        public TodoPriority Priority { get; set; }

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
            return this.Text;
        }
    }
}
