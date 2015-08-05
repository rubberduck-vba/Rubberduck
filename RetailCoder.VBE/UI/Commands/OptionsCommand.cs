using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Commands
{
    public class OptionsCommand : RubberduckCommandBase
    {
        private readonly VBE _vbe;
        private readonly IGeneralConfigService _configService;

        public OptionsCommand(IRubberduckMenuCommand command, VBE vbe, IGeneralConfigService configService)
            : base(command)
        {
            _vbe = vbe;
            _configService = configService;
        }

        public override void Initialize()
        {
            var parent = _vbe.CommandBars[1].Controls.OfType<CommandBarPopup>()
                .SingleOrDefault(control => control.Caption == RubberduckUI.RubberduckMenu);

            if (parent == null)
            {
                throw new ParentMenuNotFoundException(RubberduckUI.RubberduckMenu);
            }

            Command.AddCommandBarButton(parent.Controls, RubberduckUI.RubberduckMenu_Options, true);
        }

        public override void ExecuteAction()
        {
            using (var window = new _SettingsDialog(_configService))
            {
                window.ShowDialog();
            }
        }
    }
}