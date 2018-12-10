extern alias Office_v8;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Office_v8::Office;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office8
{
    public class CommandBarPopup : CommandBarControl, ICommandBarPopup
    {
        private readonly IVBE _vbe;

        public CommandBarPopup(MSO.CommandBarPopup target, IVBE vbe, bool rewrapping = false) 
            : base(target, vbe, rewrapping)
        {
            _vbe = vbe;
        }        

        private MSO.CommandBarPopup Popup => Target as MSO.CommandBarPopup;

        public ICommandBar CommandBar => new CommandBar(IsWrappingNullReference ? null : Popup.CommandBar, _vbe);

        public ICommandBarControls Controls => new CommandBarControls(IsWrappingNullReference ? null : Popup.Controls, _vbe);

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}