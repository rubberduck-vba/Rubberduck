using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarControls : SafeComWrapper<Microsoft.Office.Core.CommandBarControls>, ICommandBarControls
    {
        public CommandBarControls(Microsoft.Office.Core.CommandBarControls target) 
            : base(target)
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
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<ICommandBarControl>(null, o => new CommandBarControl(null))
                : new ComWrapperEnumerator<ICommandBarControl>(Target,
                    o => new CommandBarControl((Microsoft.Office.Core.CommandBarControl) o));
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


        //This does not release the children because they usually require their own cleanup, which should make the decision whather to release or not.
        private bool _isReleased;
        public virtual void Release(bool final = false)
        {
            if (IsWrappingNullReference || _isReleased || !Marshal.IsComObject(Target))
            {
                _isReleased = true;
                return;
            }

            try
            {
                if (final)
                {
                    Marshal.FinalReleaseComObject(Target);
                }
                else
                {
                    Marshal.ReleaseComObject(Target);
                }
            }
            finally
            {
                _isReleased = true;
            }
        }
    }
}