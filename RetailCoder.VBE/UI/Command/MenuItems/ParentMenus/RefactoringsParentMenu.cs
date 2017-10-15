using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class RefactoringsParentMenu : ParentMenuItemBase
    {
        public RefactoringsParentMenu(IEnumerable<IMenuItem> items)
            : base("RubberduckMenu_Refactor", items)
        {
        }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Refactorings; } }
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
