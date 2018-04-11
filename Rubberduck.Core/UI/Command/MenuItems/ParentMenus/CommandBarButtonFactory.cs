using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Rubberduck.VBEditor.Utility;

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
        private readonly List<CommandBarButton> _buttons;

        public CommandBarButtonFactory(IVBEEvents vbeEvents)
        {
            _buttons = new List<CommandBarButton>();
            _vbeEvents = vbeEvents;
            _vbeEvents.EventsTerminated += EventsTerminated;
        }

        private void EventsTerminated(object sender, System.EventArgs e)
        {
            _buttons.ForEach(b => b.DetachEvents());
        }

        public ICommandBarButton Create<TParent>(TParent parent, int? beforeIndex = null)
            where TParent : ICommandBarControls
        {
            var button = CommandBarButton.FromCommandBarControl(beforeIndex.HasValue
                ? parent.Add(ControlType.Button, beforeIndex.Value)
                : parent.Add(ControlType.Button));
            if (!button.IsWrappingNullReference)
            {
                _buttons.Add(button);
            }

            return button;
        }
    }
}