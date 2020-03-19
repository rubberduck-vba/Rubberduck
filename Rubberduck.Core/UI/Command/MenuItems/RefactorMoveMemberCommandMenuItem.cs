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

        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.MoveMember;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
