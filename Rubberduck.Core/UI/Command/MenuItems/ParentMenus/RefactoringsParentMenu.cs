using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class RefactoringsParentMenu : ParentMenuItemBase
    {
        public RefactoringsParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher)
            : base(dispatcher,"RubberduckMenu_Refactor", items)
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
        MoveToFolder,
        MoveContainingFolder,
        AddRemoveReferences
    }
}
