using System;
using System.Drawing;
using Rubberduck.Properties;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class ReparseCommandMenuItem : CommandMenuItemBase
    {
        public ReparseCommandMenuItem(CommandBase command) : base(command)
        {
        }

        private string _caption;
        public void SetCaption(string caption)
        {
            _caption = caption;
        }
        public override Func<string> Caption { get { return () => _caption; } }

        private string _tooltip;
        public void SetToolTip(string tooltip)
        {
            _tooltip = tooltip;
        }
        public override Func<string> ToolTipText { get { return () => _tooltip; } }

        public override ButtonStyle ButtonStyle => ButtonStyle.IconAndCaption;
        public override string Key => "HotkeyDescription_ParseAll";
        public override Image Image => Resources.arrow_circle_double;
        public override Image Mask => Resources.arrow_circle_double_mask;
        public override int DisplayOrder => (int)RubberduckCommandBarItemDisplayOrder.RequestReparse;
    }
}