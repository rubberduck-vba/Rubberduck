using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ParserStateCommandBar
    {
        private readonly IRubberduckParser _parser;
        private readonly VBE _vbe;

        private CommandBarButton _statusButton;

        public ParserStateCommandBar(IRubberduckParser parser, VBE vbe)
        {
            _parser = parser;
            _vbe = vbe;
            _parser.State.StateChanged += State_StateChanged;
            Initialize();
        }

        private void State_StateChanged(object sender, EventArgs e)
        {
            _statusButton.Caption = _parser.State.Status.ToString();
        }

        public void Initialize()
        {
            var commandbar = _vbe.CommandBars.Add("Parsing", MsoBarPosition.msoBarTop, false, true);
            _statusButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            _statusButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            ParentMenuItemBase.SetButtonImage(_statusButton, Resources.flask, Resources.flask_mask);
            commandbar.Visible = true;
        }
    }
}
