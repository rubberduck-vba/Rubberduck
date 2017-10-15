using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class NoIndentAnnotationCommandMenuItem : CommandMenuItemBase
    {
        public NoIndentAnnotationCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override string Key { get { return "NoIndentAnnotation"; } }
        public override int DisplayOrder { get { return (int)SmartIndenterMenuItemDisplayOrder.NoIndentAnnotation; } }
    }
}
