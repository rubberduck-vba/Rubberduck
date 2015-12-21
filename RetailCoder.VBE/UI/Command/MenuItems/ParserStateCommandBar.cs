using System;
using System.Collections.Generic;
using System.Drawing;
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

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;

        public ParserStateCommandBar(IRubberduckParser parser, VBE vbe)
        {
            _parser = parser;
            _vbe = vbe;
            _parser.State.StateChanged += State_StateChanged;
            Initialize();
        }

        //private static readonly IDictionary<ParserState, Image> ParserIcons =
        //    new Dictionary<ParserState, Image>
        //    {
        //        { ParserState.Error, Resources.balloon_prohibition },
        //        { ParserState.Resolving, Resources.balloon_ellipsis },
        //        { ParserState.Parsing, Resources.balloon_ellipsis },
        //        { ParserState.Parsed, Resources.balloon_smiley },
        //        { ParserState.Ready, Resources.balloon_smiley },
        //    };

        public void SetStatusText(string value)
        {
            _statusButton.Caption = value;
        }

        private void State_StateChanged(object sender, EventArgs e)
        {
            _statusButton.Caption = _parser.State.Status.ToString();

            // bug: apparently setting a button's icon *after* initialization blows Excel up
            //var icon = ParserIcons[_parser.State.Status];
            //ParentMenuItemBase.SetButtonImage(_statusButton, icon, Resources.balloon_mask);
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
            var commandbar = _vbe.CommandBars.Add("Parsing", MsoBarPosition.msoBarTop, false, true);

            _refreshButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            ParentMenuItemBase.SetButtonImage(_refreshButton, Resources.arrow_circle_double, Resources.arrow_circle_double_mask);
            _refreshButton.Style = MsoButtonStyle.msoButtonIcon;
            _refreshButton.TooltipText = "Parse all opened projects";
            _refreshButton.Click += refreshButton_Click;

            _statusButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            _statusButton.Style = MsoButtonStyle.msoButtonCaption;

            commandbar.Visible = true;
        }

        private void refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRefresh();
        }
    }
}
