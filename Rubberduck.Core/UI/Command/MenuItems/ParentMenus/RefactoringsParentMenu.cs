using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class RefactoringsParentMenu : ParentMenuItemBase
    {
        public RefactoringsParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items)
            : base(buttonFactory, "RubberduckMenu_Refactor", items)
        {
        }

        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.Refactorings;
    }

    public enum RefactoringsMenuItemDisplayOrder
    {
        RenameIdentifier,
        ExtractMethod,
        ExtractInterface,
        ImplementInterface,
        RemoveParameters,
        ReorderParameters,
        MoveCloserToUsage,
        EncapsulateField,
        IntroduceParameter,
        IntroduceField,
    }
}
