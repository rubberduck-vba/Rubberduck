using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Command.MenuItems.ToolStripItems
{
    public interface ICommandToolStripItem
    {
        ToolStripItem Item { get; }
        ICommand Command { get; }
        int DisplayOrder { get; }
        Func<string> Caption { get; }
        Func<string> ToolTip { get; }
    }

    public class CommandToolStripItem : ICommandToolStripItem
    {
        private readonly ToolStripItem _item;
        private readonly ICommand _command;
        private readonly int _displayOrder;
        private readonly Func<string> _caption;
        private readonly Func<string> _toolTip;

        public CommandToolStripItem(ToolStripItem item, ICommand command, int displayOrder, Func<string> caption = null, Func<string> toolTip = null)
        {
            _item = item;
            _command = command;
            _displayOrder = displayOrder;
            _caption = caption;
            _toolTip = toolTip;

            if (command != null)
            {
                item.Click += delegate { command.Execute(); };
            }
        }

        public ToolStripItem Item { get {return _item; } }
        public ICommand Command { get { return _command; } }
        public int DisplayOrder { get { return _displayOrder; } }
        public Func<string> Caption { get { return _caption; }}
        public Func<string> ToolTip { get { return _toolTip; } }
    }
}
