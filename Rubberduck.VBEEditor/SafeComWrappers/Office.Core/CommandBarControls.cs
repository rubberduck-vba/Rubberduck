using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarControls : SafeComWrapper<Microsoft.Office.Core.CommandBarControls>, ICommandBarControls
    {
        public CommandBarControls(Microsoft.Office.Core.CommandBarControls comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public ICommandBar Parent
        {
            get { return new CommandBar(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public ICommandBarControl this[object index]
        {
            get { return new CommandBarControl(ComObject[index]); }
        }

        public ICommandBarControl Add(ControlType type)
        {
            return new CommandBarControl(ComObject.Add(type, Temporary:true));
        }

        public ICommandBarControl Add(ControlType type, int before)
        {
            return new CommandBarControl(ComObject.Add(type, Before:before, Temporary:true));
        }

        IEnumerator<ICommandBarControl> IEnumerable<ICommandBarControl>.GetEnumerator()
        {
            return new ComWrapperEnumerator<CommandBarControl>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<ICommandBarControl>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Office.Core.CommandBarControls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(ICommandBarControls other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBarControls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}