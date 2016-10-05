namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBarPopup : CommandBarControl
    {
        public CommandBarPopup(Microsoft.Office.Core.CommandBarPopup comObject) 
            : base(comObject)
        {
        }

        public static CommandBarPopup FromCommandBarControl(CommandBarControl control)
        {
            return new CommandBarPopup((Microsoft.Office.Core.CommandBarPopup)control.ComObject);
        }

        private Microsoft.Office.Core.CommandBarPopup Popup
        {
            get { return (Microsoft.Office.Core.CommandBarPopup)ComObject; }
        }

        public CommandBar CommandBar
        {
            get { return new CommandBar(IsWrappingNullReference ? null : InvokeResult(() => Popup.CommandBar)); }
        }

        public CommandBarControls Controls
        {
            get { return new CommandBarControls(IsWrappingNullReference ? null : InvokeResult(() => Popup.Controls)); }
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