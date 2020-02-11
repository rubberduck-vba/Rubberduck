using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorMoveMemberCommandMenuItem : CommandMenuItemBase
    {
        public RefactorMoveMemberCommandMenuItem(RefactorMoveMemberCommand command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_MoveMember";

        //TODO: Remove Override once text is added to RubberduckMenus.resx
        public override Func<string> Caption
        {
            get
            {
                return () => "Move Member";
            }
        }
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.MoveMember;
        ////TODO: Get MoveMember Image and Mask
        //public override Image Image => Resources.CommandBarIcons.ExtractInterface;
        //public override Image Mask => Resources.CommandBarIcons.ExtractInterfaceMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
