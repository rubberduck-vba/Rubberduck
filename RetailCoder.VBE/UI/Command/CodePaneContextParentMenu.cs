using System;
using System.Collections.Generic;

namespace Rubberduck.UI.Command
{
    public class CodePaneContextParentMenu : ParentMenuItemBase
    {
        public CodePaneContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }
    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public class CodePaneContextMenuAttribute : Attribute
    {
    }
}