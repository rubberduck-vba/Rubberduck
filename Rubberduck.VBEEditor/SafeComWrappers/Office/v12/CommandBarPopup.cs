using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.v12
{
    public class CommandBarPopup : CommandBarControl, ICommandBarPopup
    {
        public CommandBarPopup(Microsoft.Office.Core.CommandBarPopup target) 
            : base(target)
        {
        }

        public static ICommandBarPopup FromCommandBarControl(ICommandBarControl control)
        {
            return new CommandBarPopup(control.Target as Microsoft.Office.Core.CommandBarPopup);
        }

        private Microsoft.Office.Core.CommandBarPopup Popup
        {
            get { return Target as Microsoft.Office.Core.CommandBarPopup; }
        }

        public ICommandBar CommandBar
        {
            get { return new CommandBar(IsWrappingNullReference ? null : Popup.CommandBar); }
        }

        public ICommandBarControls Controls
        {
            get { return new CommandBarControls(IsWrappingNullReference ? null : Popup.Controls); }
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        Controls.Release();
        //    }
        //    base.Release(final);
        //}
    }
}