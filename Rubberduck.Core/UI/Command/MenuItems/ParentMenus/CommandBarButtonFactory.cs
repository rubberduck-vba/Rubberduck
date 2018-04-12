using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public interface ICommandBarButtonFactory
    {
        ICommandBarButton Create<TParent>(TParent parent, int? beforeIndex = null)
            where TParent : ICommandBarControls;
    }

    public class CommandBarButtonFactory : ICommandBarButtonFactory
    {
        private readonly IVBEEvents _vbeEvents;
        private readonly Dictionary<int, CommandBarButton> _buttons;

        public CommandBarButtonFactory(IVBEEvents vbeEvents)
        {
            _buttons = new Dictionary<int, CommandBarButton>();
            _vbeEvents = vbeEvents;
            _vbeEvents.EventsTerminated += EventsTerminated;
        }

        private void EventsTerminated(object sender, System.EventArgs e)
        {
            foreach (var kvp in _buttons)
            {
                kvp.Value.DetachEvents();
            }
        }

        public ICommandBarButton Create<TParent>(TParent parent, int? beforeIndex = null)
            where TParent : ICommandBarControls
        {
            var button = CommandBarButton.FromCommandBarControl(beforeIndex.HasValue
                ? parent.Add(ControlType.Button, beforeIndex.Value)
                : parent.Add(ControlType.Button));
            if (!button.IsWrappingNullReference)
            {
                button.Disposing += ButtonDisposing;
                _buttons.Add(button.GetHashCode(), button);
            }

            return button;
        }

        private void ButtonDisposing(object sender, System.EventArgs e)
        {
            var button = (CommandBarButton) sender;
            button.Disposing -= ButtonDisposing;
            _buttons.Remove(button.GetHashCode());
        }
    }
}