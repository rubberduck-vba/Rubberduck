using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Input;

namespace Rubberduck.UI.Command.MenuItems.ToolStripItems
{
    public interface ICommandToolStripItem : IToolStripItem
    {
        ICommand Command { get; }
    }

    public interface IToolStripItem
    {
        ToolStripItem Item { get; }
        int DisplayOrder { get; }
        void Localize();
    }

    public class CommandToolStripItem : ICommandToolStripItem
    {
        private readonly ICommand _command;
        private readonly ToolStripItem _item;
        private readonly int _displayOrder;
        private readonly Func<string> _caption;
        private readonly Func<string> _toolTip;

        public CommandToolStripItem(ICommand command, ToolStripItem item, int displayOrder, Func<string> caption = null, Func<string> toolTip = null)
        {
            _command = command;
            _item = item;
            _displayOrder = displayOrder;
            _caption = caption;
            _toolTip = toolTip;

            if (command != null)
            {
                item.Click += delegate { command.Execute(null); };
            }
        }

        public void Localize()
        {
            if (_caption != null)
            {
                _item.Text = _caption.Invoke();
            }
            if (_toolTip != null)
            {
                _item.ToolTipText = _toolTip.Invoke();
            }
        }

        public ICommand Command { get { return _command; } }
        public ToolStripItem Item { get { return _item; } }
        public int DisplayOrder { get { return _displayOrder; } }
    }
}
