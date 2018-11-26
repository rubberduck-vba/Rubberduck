using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Microsoft.Office.Core;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office12
{
    public class CommandBarPopup : CommandBarControl, ICommandBarPopup
    {
        public CommandBarPopup(MSO.CommandBarPopup target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }        

        private MSO.CommandBarPopup Popup => Target as MSO.CommandBarPopup;

        public ICommandBar CommandBar => new CommandBar(IsWrappingNullReference ? null : Popup.CommandBar);

        public ICommandBarControls Controls => new CommandBarControls(IsWrappingNullReference ? null : Popup.Controls);

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}