using System;
using System.Collections.Generic;

namespace Rubberduck.UI.Command
{
    public class FormDesignerContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }
    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public class FormDesignerContextMenuAttribute : Attribute
    {
    }

    public class FormDesignerControlContextParentMenu : ParentMenuItemBase
    {
        public FormDesignerControlContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }
    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public class FormDesignerControlContextMenuAttribute : Attribute
    {
    }
}