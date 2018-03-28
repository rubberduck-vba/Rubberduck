﻿using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class IndentCurrentProjectCommandMenuItem : CommandMenuItemBase
    {
        public IndentCurrentProjectCommandMenuItem(CommandBase command) : base(command) { }

        public override string Key => "IndentCurrentProject";
        public override int DisplayOrder => (int)SmartIndenterMenuItemDisplayOrder.CurrentProject;
    }
}
