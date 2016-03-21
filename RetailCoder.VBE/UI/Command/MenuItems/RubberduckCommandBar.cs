using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RubberduckCommandBar
    {
        private readonly RubberduckParserState _state;
        private readonly VBE _vbe;

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;

        public RubberduckCommandBar(RubberduckParserState state, VBE vbe)
        {
            _state = state;
            _vbe = vbe;
            _state.StateChanged += State_StateChanged;
            Initialize();
        }

        public void SetStatusText(string value = null)
        {
            _statusButton.Caption = value ?? RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status);
        }

        private void State_StateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("RubberduckCommandBar handles StateChanged...");
            UiDispatcher.Invoke(() => SetStatusText(RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status)));
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
            _refreshButton.TooltipText =RubberduckUI.RubberduckCommandbarRefreshButtonTooltip;
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
