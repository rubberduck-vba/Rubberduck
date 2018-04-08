using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarPopup : CommandBarControl, ICommandBarPopup
    {
        public CommandBarPopup(Microsoft.Office.Core.CommandBarPopup target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public static ICommandBarPopup FromCommandBarControl(ICommandBarControl control)
        {
            return new CommandBarPopup(control.Target as Microsoft.Office.Core.CommandBarPopup, rewrapping: true);
        }

        private Microsoft.Office.Core.CommandBarPopup Popup => Target as Microsoft.Office.Core.CommandBarPopup;

        public ICommandBar CommandBar => new CommandBar(IsWrappingNullReference ? null : Popup.CommandBar);

        public ICommandBarControls Controls => new CommandBarControls(IsWrappingNullReference ? null : Popup.Controls);
    }
}