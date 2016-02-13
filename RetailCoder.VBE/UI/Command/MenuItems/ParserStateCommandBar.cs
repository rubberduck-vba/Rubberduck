using System;
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

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;

        public ParserStateCommandBar(IRubberduckParser parser, VBE vbe)
        {
            _parser = parser;
            _vbe = vbe;
            _parser.State.StateChanged += State_StateChanged;
            Initialize();
        }

        public void SetStatusText(string value)
        {
            _statusButton.Caption = value;
        }

        private void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            UiDispatcher.Invoke(() => _statusButton.Caption = e.State.ToString());
        }

        public event EventHandler Refresh;

        private void OnRefresh()
        {
            var handler = Refresh;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        public void Initialize()
        {
            var commandbar = _vbe.CommandBars.Add("Rubberduck", MsoBarPosition.msoBarTop, false, true);

            _refreshButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            ParentMenuItemBase.SetButtonImage(_refreshButton, Resources.arrow_circle_double, Resources.arrow_circle_double_mask);
            _refreshButton.Style = MsoButtonStyle.msoButtonIcon;
            _refreshButton.Tag = "Refresh";
            _refreshButton.TooltipText = "Parse all opened projects";
            _refreshButton.Click += refreshButton_Click;

            _statusButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            _statusButton.Style = MsoButtonStyle.msoButtonCaption;
            _statusButton.Tag = "Status";

            commandbar.Visible = true;
        }

        private void refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRefresh();
        }
    }
}
