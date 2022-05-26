using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class CodePaneContextParentMenu : ParentMenuItemBase
    {
        public CodePaneContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex, IUiDispatcher dispatcher)
            : base(dispatcher,"RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public enum CodePaneContextMenuItemDisplayOrder
    {
        Refactorings,
        Annotate,
        Indenter,
        RegexSearchReplace,
        PeekDefinition,
        FindSymbol,
        FindAllReferences,
        FindAllImplementations,
        RunSelectedTestModule,
        RunSelectedTest
    }
}
