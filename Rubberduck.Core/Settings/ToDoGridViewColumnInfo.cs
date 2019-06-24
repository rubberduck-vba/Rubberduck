using System;
using Rubberduck.UI;
using System.Windows.Controls;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IGridViewColumnInfo
    {
        int DisplayIndex { get; set; }
        DataGridLength Width { get; set; }
    }

    public class ToDoGridViewColumnInfo : ViewModelBase, IGridViewColumnInfo, IEquatable<ToDoGridViewColumnInfo>
    {
        private int _displayIndex;
        public int DisplayIndex
        {
            get => _displayIndex;
            set
            {
                if (value != _displayIndex)
                {
                    _displayIndex = value;
                    OnPropertyChanged();
                }
            }
        }

        [XmlElement(Type = typeof(DataGridLength))]
        private DataGridLength _width;

        public DataGridLength Width
        {
            get => _width;
            set
            {
                if (value != _width)
                {
                    _width = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ToDoGridViewColumnInfo()
        {
        }

        public ToDoGridViewColumnInfo(int displayIndex, DataGridLength width)
        {
            DisplayIndex = displayIndex;
            Width = width;
        }

        public bool Equals(ToDoGridViewColumnInfo other)
        {
            return DisplayIndex == other.DisplayIndex
                && Width == other.Width;
        }
    }
}
