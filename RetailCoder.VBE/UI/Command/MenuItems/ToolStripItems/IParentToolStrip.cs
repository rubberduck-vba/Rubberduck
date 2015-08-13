using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.UI.Command.MenuItems.ToolStripItems
{
    public interface IParentToolStrip
    {
        void Localize();
        void Initialize(ToolStrip toolStrip);
    }

    public abstract class ParentToolStripBase : IParentToolStrip
    {
        private readonly IDictionary<ICommandToolStripItem, ToolStripItem> _items;

        protected ParentToolStripBase(IEnumerable<ICommandToolStripItem> items)
        {
            _items = items.ToDictionary(item => item, item => item.Item);
        }

        public void Localize()
        {
            foreach (var kvp in _items)
            {
                kvp.Value.Text = kvp.Key.Caption.Invoke();
                kvp.Value.ToolTipText = kvp.Key.ToolTip.Invoke();
            }
        }

        private ToolStrip _toolStrip;
        public void Initialize(ToolStrip toolStrip)
        {
            if (_toolStrip != null)
            {
                throw new InvalidOperationException("Instance is already initialized.");
            }
            
            var items = _items
                .OrderBy(item => item.Key.DisplayOrder)
                .Select(item => item.Value)
                .ToArray();

            _toolStrip = toolStrip;
            _toolStrip.Items.AddRange(items);

            Localize();
        }
    }
}