using System;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.Commands
{
    public class AboutCommand : ICommand
    {
        public void Execute()
        {
            using (var window = new _AboutWindow())
            {
                window.ShowDialog();
            }
        }
    }

    public class AboutCommandMenuItem : CommandMenuItemBase
    {
        public AboutCommandMenuItem(ICommand command)
            : base(command)
        {
        }
        
        public override string Key { get { return "RubberduckMenu_About"; } }
    }
}