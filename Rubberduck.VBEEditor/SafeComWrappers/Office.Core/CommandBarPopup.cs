using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarPopup : CommandBarControl, ICommandBarPopup
    {
        public CommandBarPopup(Microsoft.Office.Core.CommandBarPopup comObject) 
            : base(comObject)
        {
        }

        public static ICommandBarPopup FromCommandBarControl(ICommandBarControl control)
        {
            return new CommandBarPopup((Microsoft.Office.Core.CommandBarPopup)control.ComObject);
        }

        private Microsoft.Office.Core.CommandBarPopup Popup
        {
            get { return (Microsoft.Office.Core.CommandBarPopup)ComObject; }
        }

        public ICommandBar CommandBar
        {
            get { return new CommandBar(IsWrappingNullReference ? null : Popup.CommandBar); }
        }

        public ICommandBarControls Controls
        {
            get { return new CommandBarControls(IsWrappingNullReference ? null : Popup.Controls); }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Controls.Release();
            }
            base.Release();
        }
    }
}