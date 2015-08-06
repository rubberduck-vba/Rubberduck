using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Command
{
    public class AboutCommand : ICommand
    {
        public void Execute()
        {
            using (var window = new AboutWindow())
            {
                window.ShowDialog();
            }
        }
    }

    public class AboutMenuCommand : CommandMenuItemBase
    {
        public AboutMenuCommand(ICommand command) : base(command)
        {
        }

        public override string Key
        {
            get { return RubberduckUI.RubberduckMenu_About; }
        }

        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.About; } }
    }
}
