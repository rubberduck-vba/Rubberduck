extern alias Office_v8;

using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Office_v8::Office;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office8
{
    public sealed class CommandBarControls : SafeComWrapper<MSO.CommandBarControls>, ICommandBarControls
    {
        private readonly IVBE _vbe;

        public CommandBarControls(MSO.CommandBarControls target, IVBE vbe, bool rewrapping = false) 
            : base(target, rewrapping)
        {
            _vbe = vbe;
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public ICommandBar Parent => new CommandBar(IsWrappingNullReference ? null : Target.Parent, _vbe);

        public ICommandBarControl this[object index] => new CommandBarControl(!IsWrappingNullReference ? Target[index] : null, _vbe);

        public ICommandBarButton AddButton(int? before = null)
        {
            return before.HasValue
                ? new CommandBarButton((IsWrappingNullReference ? null : Target.Add(ControlType.Button, Before: before, Temporary: CommandBarControl.AddCommandBarControlsTemporarily) as MSO.CommandBarButton), _vbe)
                : new CommandBarButton((IsWrappingNullReference ? null : Target.Add(ControlType.Button, Temporary: CommandBarControl.AddCommandBarControlsTemporarily) as MSO.CommandBarButton), _vbe);            
        }

        public ICommandBarPopup AddPopup(int? before = null)
        {
            return before.HasValue
                ? new CommandBarPopup(IsWrappingNullReference ? null : Target.Add(ControlType.Popup, Before: before, Temporary: CommandBarControl.AddCommandBarControlsTemporarily) as MSO.CommandBarPopup, _vbe)
                : new CommandBarPopup(IsWrappingNullReference ? null : Target.Add(ControlType.Popup, Temporary: CommandBarControl.AddCommandBarControlsTemporarily) as MSO.CommandBarPopup, _vbe);
        }

        IEnumerator<ICommandBarControl> IEnumerable<ICommandBarControl>.GetEnumerator()
        {
            return new ComWrapperEnumerator<ICommandBarControl>(Target,
                    comObject => new CommandBarControl((MSO.CommandBarControl) comObject, _vbe));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new List<ICommandBarControl>().GetEnumerator()
                : ((IEnumerable<ICommandBarControl>) this).GetEnumerator();
        }

        public override bool Equals(ISafeComWrapper<MSO.CommandBarControls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICommandBarControls other)
        {
            return Equals(other as SafeComWrapper<MSO.CommandBarControls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}