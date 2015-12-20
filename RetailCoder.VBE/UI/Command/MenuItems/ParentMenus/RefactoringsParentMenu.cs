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
        ExtractMethod,
        RenameIdentifier,
        ReorderParameters,
        RemoveParameters,
        IntroduceParameter,
        IntroduceField,
        EncapsulateField,
    }
}