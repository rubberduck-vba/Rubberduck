using System;
using System.Collections.Generic;

namespace Rubberduck.UI.Command
{
    public class ProjectWindowContextParentMenu : ParentMenuItemBase
    {
        public ProjectWindowContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }        
    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public class ProjectWindowContextMenuAttribute : Attribute
    {
    }
}