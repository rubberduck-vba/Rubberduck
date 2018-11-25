using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public interface IAppCommandBar : IAppMenu
    {
        ICommandBars Parent { get; set; }
        ICommandBar Item { get; }
        void RemoveCommandBar();
    }
}