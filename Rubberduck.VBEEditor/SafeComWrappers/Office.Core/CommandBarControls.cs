using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarControls : SafeComWrapper<Microsoft.Office.Core.CommandBarControls>, ICommandBarControls
    {
        public CommandBarControls(Microsoft.Office.Core.CommandBarControls target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public ICommandBar Parent => new CommandBar(IsWrappingNullReference ? null : Target.Parent);

        public ICommandBarControl this[object index] => new CommandBarControl(!IsWrappingNullReference ? Target[index] : null);

        public ICommandBarControl Add(ControlType type)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.Add(type, Temporary: CommandBarControl.AddCommandBarControlsTemporarily));
        }

        public ICommandBarControl Add(ControlType type, int before)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.Add(type, Before: before, Temporary: CommandBarControl.AddCommandBarControlsTemporarily));
        }

        IEnumerator<ICommandBarControl> IEnumerable<ICommandBarControl>.GetEnumerator()
        {
            return new ComWrapperEnumerator<ICommandBarControl>(Target,
                    comObject => new CommandBarControl((Microsoft.Office.Core.CommandBarControl) comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new List<ICommandBarControl>().GetEnumerator()
                : ((IEnumerable<ICommandBarControl>) this).GetEnumerator();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(ISafeComWrapper<Microsoft.Office.Core.CommandBarControls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICommandBarControls other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBarControls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}