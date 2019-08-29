using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class NoIndentAnnotationCommandMenuItem : CommandMenuItemBase
    {
        public NoIndentAnnotationCommandMenuItem(NoIndentAnnotationCommand command)
            : base(command)
        {
        }

        public override string Key => "NoIndentAnnotation";
        public override int DisplayOrder => (int)SmartIndenterMenuItemDisplayOrder.NoIndentAnnotation;
    }
}
