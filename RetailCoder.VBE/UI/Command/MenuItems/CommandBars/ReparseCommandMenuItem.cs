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

        public override ButtonStyle ButtonStyle { get { return ButtonStyle.Icon; } }

        public override string Key { get { return "HotkeyDescription_ParseAll"; } }
        public override Image Image { get { return Resources.arrow_circle_double; } }
        public override Image Mask { get { return Resources.arrow_circle_double_mask; } }
        public override int DisplayOrder { get { return (int)RubberduckCommandBarItemDisplayOrder.RequestReparse; } }
    }
}