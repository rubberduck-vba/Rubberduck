using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public interface IAppCommandBar : IAppMenu
    {
        ICommandBars Parent { get; set; }
        ICommandBar Item { get; }
        void RemoveChildren();
    }
}