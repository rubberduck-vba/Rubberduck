﻿using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class SourceControlCommandMenuItem : CommandMenuItemBase
    {
        public SourceControlCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key => "ToolsMenu_SourceControl";
        public override int DisplayOrder => (int)ToolsMenuItemDisplayOrder.SourceControl;
    }
}
