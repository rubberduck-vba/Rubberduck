using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.Commands
{
    /// <summary>
    /// A command that displays the "About" dialog.
    /// </summary>
    public class AboutCommand : RubberduckCommandBase
    {
        private readonly VBE _vbe;

        public AboutCommand(IRubberduckMenuCommand command, VBE vbe)
            : base(command)
        {
            _vbe = vbe;
        }

        public override void Initialize()
        {
            var parent = _vbe.CommandBars[1].Controls.OfType<CommandBarPopup>()
                .SingleOrDefault(control => control.Caption == RubberduckUI.RubberduckMenu);

            if (parent == null)
            {
                throw new InvalidOperationException("Parent menu not found. Cannot create child menu item.");
            }

            Command.AddCommandBarButton(parent.Controls, RubberduckUI.RubberduckMenu_About, true);
        }

        public override void ExecuteAction()
        {
            using (var window = new _AboutWindow())
            {
                window.ShowDialog();
            }
        }
    }
}