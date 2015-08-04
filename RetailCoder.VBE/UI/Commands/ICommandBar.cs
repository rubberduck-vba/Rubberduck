using System.Collections.Generic;

namespace Rubberduck.UI.Commands
{
    public interface ICommandBar
    {
        void Localize();
        void AddItem(IMenuItem item, bool? beginGroup = null, int? beforeIndex = null);
        bool RemoveItem(IMenuItem item);
        bool Remove();
        IEnumerable<IMenuItem> Items { get; }
    }
}